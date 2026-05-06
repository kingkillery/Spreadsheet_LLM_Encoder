"""Tests for llm_backend.py — EchoBackend, CallableBackend, OpenAIBackend."""
import os
import sys
import types
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.dirname(__file__))
from llm_backend import CallableBackend, EchoBackend, LLMBackend, OpenAIBackend


class TestEchoBackendSingleResponse(unittest.TestCase):

    def setUp(self):
        self.backend = EchoBackend(response="x")

    def test_returns_configured_response(self):
        self.assertEqual(self.backend("any prompt"), "x")

    def test_records_all_calls(self):
        self.backend("first")
        self.backend("second")
        self.backend("third")
        self.assertEqual(self.backend.calls, ["first", "second", "third"])

    def test_always_returns_same_response_regardless_of_prompt(self):
        responses = {self.backend(f"prompt {i}") for i in range(5)}
        self.assertEqual(responses, {"x"})


class TestEchoBackendResponseCycling(unittest.TestCase):

    def setUp(self):
        self.backend = EchoBackend(responses=["a", "b"])

    def test_cycles_through_responses(self):
        self.assertEqual(self.backend("p1"), "a")
        self.assertEqual(self.backend("p2"), "b")
        self.assertEqual(self.backend("p3"), "a")
        self.assertEqual(self.backend("p4"), "b")

    def test_all_calls_recorded_during_cycling(self):
        for i in range(4):
            self.backend(f"call {i}")
        self.assertEqual(len(self.backend.calls), 4)


class TestEchoBackendDefaultResponse(unittest.TestCase):

    def test_default_response_is_empty_string(self):
        backend = EchoBackend()
        self.assertEqual(backend("prompt"), "")


class TestEchoBackendSatisfiesProtocol(unittest.TestCase):

    def test_echo_backend_is_llm_backend(self):
        self.assertIsInstance(EchoBackend(), LLMBackend)


class TestCallableBackend(unittest.TestCase):

    def test_delegates_to_wrapped_callable(self):
        backend = CallableBackend(lambda p: p.upper())
        self.assertEqual(backend("hello"), "HELLO")

    def test_callable_backend_satisfies_protocol(self):
        self.assertIsInstance(CallableBackend(str), LLMBackend)

    def test_callable_receives_exact_prompt(self):
        received = []
        backend = CallableBackend(lambda p: received.append(p) or "ok")
        backend("my prompt")
        self.assertEqual(received, ["my prompt"])


class TestOpenAIBackendConstructor(unittest.TestCase):

    def test_default_model_is_gpt_4o_mini(self):
        backend = OpenAIBackend()
        self.assertEqual(backend.model, "gpt-4o-mini")

    def test_custom_model_stored(self):
        backend = OpenAIBackend(model="gpt-4")
        self.assertEqual(backend.model, "gpt-4")

    def test_system_prompt_stored(self):
        backend = OpenAIBackend(system_prompt="You are a helpful assistant.")
        self.assertEqual(backend.system_prompt, "You are a helpful assistant.")

    def test_injected_client_stored(self):
        mock_client = MagicMock()
        backend = OpenAIBackend(client=mock_client)
        self.assertIs(backend._client, mock_client)


class TestOpenAIBackendMissingPackage(unittest.TestCase):

    def test_ensure_client_raises_runtime_error_when_openai_missing(self):
        backend = OpenAIBackend()  # no client injected
        # Simulate openai not being installed by patching the import
        with patch.dict(sys.modules, {"openai": None}):
            with self.assertRaises((RuntimeError, ImportError)):
                backend._ensure_client()


class TestOpenAIBackendCallWithMockClient(unittest.TestCase):

    def _make_mock_response(self, content: str):
        choice = MagicMock()
        choice.message.content = content
        resp = MagicMock()
        resp.choices = [choice]
        return resp

    def test_call_invokes_chat_completions_create(self):
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = self._make_mock_response("answer")
        backend = OpenAIBackend(model="gpt-4o-mini", client=mock_client)
        result = backend("What is 2+2?")
        mock_client.chat.completions.create.assert_called_once()
        self.assertEqual(result, "answer")

    def test_call_passes_model_to_create(self):
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = self._make_mock_response("ok")
        backend = OpenAIBackend(model="gpt-4", client=mock_client)
        backend("prompt")
        call_kwargs = mock_client.chat.completions.create.call_args
        self.assertEqual(call_kwargs.kwargs.get("model") or call_kwargs.args[0], "gpt-4")

    def test_call_includes_system_prompt_when_set(self):
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = self._make_mock_response("ok")
        backend = OpenAIBackend(system_prompt="Be concise.", client=mock_client)
        backend("hello")
        messages = mock_client.chat.completions.create.call_args.kwargs["messages"]
        self.assertEqual(messages[0], {"role": "system", "content": "Be concise."})

    def test_call_omits_system_message_when_none(self):
        mock_client = MagicMock()
        mock_client.chat.completions.create.return_value = self._make_mock_response("ok")
        backend = OpenAIBackend(system_prompt=None, client=mock_client)
        backend("hello")
        messages = mock_client.chat.completions.create.call_args.kwargs["messages"]
        roles = [m["role"] for m in messages]
        self.assertNotIn("system", roles)


if __name__ == "__main__":
    unittest.main()
