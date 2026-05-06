# Fine-tuning SpreadsheetLLM table-detection models

## Overview

SpreadsheetLLM enables fine-tuning both closed-model (GPT-3.5/4) and open-source (Llama 2/3, Phi-2/3, Mistral) language models on the table-detection task. The approach encodes spreadsheets using the SheetCompressor to produce compact prompt-completion pairs suitable for supervised fine-tuning, following the methodology described in the paper (arXiv:2407.09025).

The fine-tuned models predict table boundaries within spreadsheets, trained on pairs of:
- **Prompt**: The table-detection instruction + the compressed encoding of the spreadsheet
- **Completion**: The ground-truth table range in Excel notation (e.g., `['range': 'A2:D5']`)

This guide covers data preparation, training via ml-intern, local LoRA fine-tuning, and integration with the Chain-of-Spreadsheet pipeline.

## Step 1 — Prepare data

### Generate fine-tuning JSONL

Use `prepare_finetuning_data.py` to convert a dataset of spreadsheets and annotations into the JSONL format required for fine-tuning:

```bash
python prepare_finetuning_data.py datasets/my_benchmark/ output/finetune.jsonl --k 4
```

**Parameters**:
- `dataset_dir`: Path to a directory containing `.xlsx` files paired with `.json` annotation files
- `output_path`: Output JSONL file path
- `--k`: Neighborhood distance for structural anchors (default: 4, paper's best ablation)

**Annotation format**: Each `.json` file must contain a `"tables"` array with table metadata:

```json
{
  "tables": [
    {"range": "A1:D5"},
    {"range": "A8:D12"}
  ]
}
```

The range format is Excel-style (`A1:D5` means columns A–D, rows 1–5).

### (Optional) Push to HuggingFace Hub

To upload the fine-tuning dataset to HuggingFace for reproducibility and team sharing:

```bash
python prepare_finetuning_data.py datasets/my_benchmark/ output/finetune.jsonl --k 4 \
  --push-to-hub username/spreadsheetllm-tabledet
```

This requires the `datasets` package (`pip install datasets`) and a HuggingFace account with `HF_TOKEN` set.

### Generate synthetic data (optional)

If you don't have a real dataset, generate synthetic spreadsheets and annotations using `scripts/synthesize_benchmark.py`:

```bash
python scripts/synthesize_benchmark.py datasets/synthetic_v1/ --n 50 --seed 42
```

See [datasets/README.md](datasets/README.md) for details on synthetic data and real dataset candidates.

## Step 2 — Fine-tune via ml-intern

**ml-intern** is HuggingFace's autonomous ML coding agent. It can orchestrate fine-tuning runs on HuggingFace compute, handling data loading, hyperparameter selection, and model pushes automatically.

### Install ml-intern

```bash
git clone https://github.com/huggingface/ml-intern.git
cd ml-intern
uv sync
uv tool install -e .
```

**Requirements**:
- `HF_TOKEN` (HuggingFace API token, for dataset/model access and HF Jobs)
- `ANTHROPIC_API_KEY` or `OPENAI_API_KEY` (for the agent's reasoning)
- `GITHUB_TOKEN` (optional, for GitHub integration)

### Run the agent

Start ml-intern in interactive mode:

```bash
ml-intern
```

Or use headless mode with a specific prompt:

```bash
ml-intern "your prompt here"
```

### Example prompt for table-detection fine-tuning

Paste this into ml-intern:

```
Fine-tune an open-source model on the table-detection task using the SheetCompresssor encoding.

Dataset: username/spreadsheetllm-tabledet (on HuggingFace Hub)
Recommended base models (from the SpreadsheetLLM paper):
- meta-llama/Llama-2-7b-hf
- meta-llama/Meta-Llama-3-8B
- microsoft/phi-2
- microsoft/Phi-3-mini-4k-instruct
- mistralai/Mistral-7B-v0.1

Fine-tuning approach:
- Method: LoRA (rank=16, alpha=32, target attention projections)
- Epochs: 3
- Learning rate: 2e-4
- Precision: fp16

Evaluation:
- Hold out 10% as validation
- Report token-level loss and sample completions

Output:
- Push the LoRA adapter to HuggingFace Hub as username/spreadsheetllm-{base_model_short}-lora
- Use the HuggingFace Jobs tool to launch the training run on HF compute

Follow the paper's methodology (arXiv:2407.09025, Section 4.1 and Appendix M.1).
```

The agent will:
1. Inspect the dataset on HuggingFace Hub
2. Select a base model
3. Configure LoRA parameters
4. Launch a training job on HF compute (via the Jobs tool)
5. Push the trained adapter to HuggingFace

## Step 3 — Fine-tune locally with LoRA

For prototyping or when HF compute is unavailable, use `scripts/finetune_local.py`:

```bash
python scripts/finetune_local.py \
  --jsonl output/finetune.jsonl \
  --base-model microsoft/phi-2 \
  --output-dir runs/phi2-sllm \
  --epochs 3
```

**Requirements**:
- CUDA-capable GPU (24GB+ VRAM recommended for 7B models)
- PyTorch with CUDA support
- `peft` and `transformers` libraries

**GPU memory ballpark**:
- 7B models + LoRA (rank=16): ~20–24GB
- 8B Llama-3 + LoRA: 24GB; use QLoRA (quantized LoRA) on smaller cards (8–16GB)

**Output**: LoRA adapter saved to `runs/phi2-sllm/adapter_model`. Push manually to HuggingFace if needed:

```bash
python -c "
from peft import PeftModel
import transformers

model = transformers.AutoModelForCausalLM.from_pretrained('microsoft/phi-2')
model = PeftModel.from_pretrained(model, 'runs/phi2-sllm')
model.push_to_hub('username/spreadsheetllm-phi2-lora')
"
```

## Step 4 — Evaluate the fine-tuned model

Integrate a fine-tuned adapter into the Chain-of-Spreadsheet pipeline by implementing a custom `LLMBackend`:

```python
from peft import PeftModel
from transformers import AutoModelForCausalLM, AutoTokenizer
import torch

class LoRABackend:
    def __init__(self, base_model: str, adapter_repo: str):
        self.model = AutoModelForCausalLM.from_pretrained(
            base_model,
            torch_dtype=torch.float16,
            device_map="auto"
        )
        self.model = PeftModel.from_pretrained(self.model, adapter_repo)
        self.tokenizer = AutoTokenizer.from_pretrained(base_model)

    def __call__(self, prompt: str) -> str:
        inputs = self.tokenizer(prompt, return_tensors="pt")
        with torch.no_grad():
            outputs = self.model.generate(**inputs, max_new_tokens=100)
        return self.tokenizer.decode(outputs[0], skip_special_tokens=True)

# Use in Chain-of-Spreadsheet
import chain_of_spreadsheet as cos

backend = LoRABackend(
    base_model="microsoft/phi-2",
    adapter_repo="username/spreadsheetllm-phi2-lora"
)
cos.configure_backend(backend)

# Now use the pipeline
from evaluation import load_spreadsheet_dataset
data = load_spreadsheet_dataset("datasets/test/")
for item in data:
    pred_tables = cos.identify_table(item["spreadsheet_path"], "Sheet1", "What tables are here?")
    print(pred_tables)
```

Then evaluate predictions using `evaluation.py`:

```python
from evaluation import load_spreadsheet_dataset, evaluate_detections

data = load_spreadsheet_dataset("datasets/test/")
all_pred_boxes = []  # Collect predictions from your fine-tuned model
all_gt_boxes = []

for item in data:
    gt_boxes = item["bboxes"]
    # Run your model to get pred_boxes...
    all_gt_boxes.extend(gt_boxes)
    all_pred_boxes.extend(pred_boxes)

precision, recall, f1 = evaluate_detections(all_pred_boxes, all_gt_boxes, threshold=0.0)
print(f"Precision: {precision:.2f}, Recall: {recall:.2f}, F1: {f1:.2f}")
```

## Tips & troubleshooting

### Distribution consistency

Keep the encoding parameters consistent across train and eval:
- Use the same `--k` (neighborhood distance) when encoding
- Use the same tokenizer model (see `Spreadsheet_LLM_Encoder.py --tokenizer-model`) across train/test splits
- Mismatches introduce distribution drift and hurt generalization

### Memory and speed

- **LoRA rank**: Start with rank=8 for faster iteration; use rank=16 for better quality
- **Batch size**: 8 for 7B + LoRA on 24GB; reduce to 4 or use gradient checkpointing for 8B models
- **Validation split**: 10% default; adjust based on dataset size (at least 5 samples recommended)

### Cost estimates

Fine-tuning via HuggingFace Jobs scales with job duration and GPU type:
- 7B model, 3 epochs on A100 40GB: ~$5–15 USD
- Larger models or more epochs scale linearly

(OpenAI fine-tuning API support is not yet integrated; see [README.md](README.md) "Not yet implemented" section.)

### Verification

After fine-tuning, test with a small sample to confirm the model learned:

```bash
python scripts/synthesize_benchmark.py datasets/test_sample/ --n 2
python scripts/finetune_local.py --jsonl datasets/test_sample/finetune.jsonl --base-model microsoft/phi-2 --epochs 1 --output-dir runs/quick_test
```

Monitor the training loss; if it plateaus early, increase epochs or check data quality.
