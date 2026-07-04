"""Unit tests for the pure logic in streamlit_app.

These do not spin up the Streamlit runtime; they exercise the testable
helpers directly so CI stays fast and deterministic.
"""

from streamlit_app import greeting


def test_greeting_with_name():
    assert greeting("Aishah") == "Hello, Aishah!"


def test_greeting_strips_whitespace():
    assert greeting("  Sam  ") == "Hello, Sam!"


def test_greeting_defaults_to_world_when_empty():
    assert greeting("") == "Hello, world!"
    assert greeting("   ") == "Hello, world!"
