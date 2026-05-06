"""
Local LoRA fine-tuning script for SpreadsheetLLM table-detection data.

Targets instruction-following and causal-LM models from the Llama/Phi/Mistral
families cited in the SpreadsheetLLM paper.  Produces a PEFT LoRA adapter that
can be merged or served directly.

Recommended --base-model values (per SpreadsheetLLM paper):
    meta-llama/Llama-2-7b-hf
    meta-llama/Meta-Llama-3-8B
    microsoft/phi-2
    microsoft/Phi-3-mini-4k-instruct
    mistralai/Mistral-7B-v0.1

Authentication for gated models:
    Set HF_TOKEN env var or run `huggingface-cli login` before calling this
    script.

Usage:
    python scripts/finetune_local.py \\
        --jsonl data/finetune.jsonl \\
        --base-model meta-llama/Meta-Llama-3-8B \\
        --output-dir runs/llama3-lora \\
        [--lora-r 16] [--lora-alpha 32] [--lr 2e-4] [--epochs 3] \\
        [--batch-size 1] [--max-seq-len 4096] \\
        [--push-to-hub username/my-adapter] [--dry-run]
"""

import argparse
import json
import logging
import sys
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Optional heavy imports — deferred so --dry-run and --help work without them.
# ---------------------------------------------------------------------------
_MISSING_DEPS: list[str] = []


def _check_deps() -> None:
    """Raise ImportError listing all missing optional dependencies at once."""
    missing = []
    for mod in ("torch", "datasets", "transformers", "peft", "accelerate"):
        try:
            __import__(mod)
        except ImportError:
            missing.append(mod)
    if missing:
        raise ImportError(
            "Missing required packages: "
            + ", ".join(missing)
            + "\nInstall with: pip install transformers peft accelerate datasets torch"
        )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _load_jsonl(path: str) -> list[dict]:
    records = []
    with open(path, encoding="utf-8") as fh:
        for lineno, raw in enumerate(fh, 1):
            raw = raw.strip()
            if not raw:
                continue
            try:
                records.append(json.loads(raw))
            except json.JSONDecodeError as exc:
                raise ValueError(f"Invalid JSON on line {lineno} of {path}: {exc}") from exc
    if not records:
        raise ValueError(f"No records found in {path}")
    logger.info("Loaded %d records from %s", len(records), path)
    return records


def _format_record(record: dict) -> str:
    """Convert a prompt/completion dict to a single training string."""
    prompt = record.get("prompt", "")
    completion = record.get("completion", "")
    return f"{prompt}\n\n{completion}"


def _build_dataset(records: list[dict], tokenizer, max_seq_len: int):
    """Tokenize records and return a datasets.Dataset ready for Trainer."""
    import datasets as _datasets  # noqa: PLC0415

    texts = [_format_record(r) for r in records]

    def _tokenize(batch):
        return tokenizer(
            batch["text"],
            truncation=True,
            max_length=max_seq_len,
            padding=False,
        )

    raw_ds = _datasets.Dataset.from_dict({"text": texts})
    tokenized = raw_ds.map(_tokenize, batched=True, remove_columns=["text"])
    tokenized = tokenized.map(
        lambda ex: {"labels": ex["input_ids"].copy()},
        batched=False,
    )
    return tokenized


def _detect_device() -> str:
    import torch  # noqa: PLC0415

    if torch.cuda.is_available():
        logger.info("CUDA device detected: %s", torch.cuda.get_device_name(0))
        return "cuda"
    if torch.backends.mps.is_available():
        logger.info("Apple MPS device detected.")
        return "mps"
    logger.warning(
        "No GPU detected — training on CPU. This will be very slow for large models."
    )
    return "cpu"


# ---------------------------------------------------------------------------
# Core training function
# ---------------------------------------------------------------------------

def train(args: argparse.Namespace) -> None:
    _check_deps()

    import torch  # noqa: PLC0415
    from datasets import Dataset  # noqa: PLC0415, F401
    from transformers import (  # noqa: PLC0415
        AutoModelForCausalLM,
        AutoTokenizer,
        DataCollatorForLanguageModeling,
        Trainer,
        TrainingArguments,
    )
    from peft import LoraConfig, TaskType, get_peft_model  # noqa: PLC0415

    device = _detect_device()
    use_fp16 = device == "cuda"

    # 1. Load tokenizer
    logger.info("Loading tokenizer from %s", args.base_model)
    tokenizer = AutoTokenizer.from_pretrained(args.base_model, use_fast=True)
    if tokenizer.pad_token is None:
        tokenizer.pad_token = tokenizer.eos_token

    # 2. Load base model
    logger.info("Loading base model from %s", args.base_model)
    model = AutoModelForCausalLM.from_pretrained(
        args.base_model,
        torch_dtype=torch.float16 if use_fp16 else torch.float32,
        device_map="auto" if device == "cuda" else None,
    )

    # 3. Wrap with LoRA
    lora_config = LoraConfig(
        task_type=TaskType.CAUSAL_LM,
        r=args.lora_r,
        lora_alpha=args.lora_alpha,
        target_modules=["q_proj", "v_proj", "k_proj", "o_proj"],
        lora_dropout=0.05,
        bias="none",
    )
    model = get_peft_model(model, lora_config)
    model.print_trainable_parameters()

    if use_fp16:
        model.enable_input_require_grads()  # needed for gradient checkpointing + fp16

    # 4. Load & tokenize data
    records = _load_jsonl(args.jsonl)
    train_ds = _build_dataset(records, tokenizer, args.max_seq_len)

    # 5. Training arguments
    output_dir = Path(args.output_dir)
    training_args = TrainingArguments(
        output_dir=str(output_dir),
        num_train_epochs=args.epochs,
        per_device_train_batch_size=args.batch_size,
        learning_rate=args.lr,
        fp16=use_fp16,
        gradient_checkpointing=use_fp16,
        save_strategy="epoch",
        logging_steps=10,
        remove_unused_columns=False,
        report_to="none",
    )

    # 6. Trainer
    data_collator = DataCollatorForLanguageModeling(tokenizer=tokenizer, mlm=False)
    trainer = Trainer(
        model=model,
        args=training_args,
        train_dataset=train_ds,
        data_collator=data_collator,
    )

    logger.info("Starting training ...")
    trainer.train()

    # 7. Save adapter
    logger.info("Saving LoRA adapter to %s", output_dir)
    model.save_pretrained(str(output_dir))
    tokenizer.save_pretrained(str(output_dir))

    # 8. Optional Hub push
    if args.push_to_hub:
        logger.info("Pushing adapter to Hub repo: %s", args.push_to_hub)
        model.push_to_hub(args.push_to_hub)
        tokenizer.push_to_hub(args.push_to_hub)
        logger.info(
            "Pushed to https://huggingface.co/%s", args.push_to_hub
        )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--jsonl",
        required=True,
        help="Path to fine-tuning JSONL (output of prepare_finetuning_data.py).",
    )
    parser.add_argument(
        "--base-model",
        required=True,
        metavar="HF_REPO",
        help="HuggingFace model repo to fine-tune (see docstring for recommendations).",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        metavar="PATH",
        help="Directory to save the LoRA adapter and tokenizer.",
    )
    parser.add_argument("--lora-r", type=int, default=16, help="LoRA rank (default: 16).")
    parser.add_argument(
        "--lora-alpha", type=int, default=32, help="LoRA alpha scaling (default: 32)."
    )
    parser.add_argument("--lr", type=float, default=2e-4, help="Learning rate (default: 2e-4).")
    parser.add_argument("--epochs", type=int, default=3, help="Training epochs (default: 3).")
    parser.add_argument(
        "--batch-size", type=int, default=1, help="Per-device train batch size (default: 1)."
    )
    parser.add_argument(
        "--max-seq-len",
        type=int,
        default=4096,
        help="Maximum token sequence length (default: 4096).",
    )
    parser.add_argument(
        "--push-to-hub",
        metavar="REPO_ID",
        default=None,
        help=(
            "HuggingFace Hub repo to push the trained adapter to. "
            "Authenticate via HF_TOKEN or `huggingface-cli login`."
        ),
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print resolved config and exit without training.",
    )
    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    if args.dry_run:
        logger.info("Dry-run mode — config:")
        for key, val in sorted(vars(args).items()):
            logger.info("  %-20s %s", key, val)
        logger.info("Exiting without training.")
        sys.exit(0)

    train(args)


if __name__ == "__main__":
    main()
