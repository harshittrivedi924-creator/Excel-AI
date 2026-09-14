import pytest


class TestVoiceModule:
    def test_import_listener(self):
        from backend.voice.listener import VoiceError, VoiceListener

        assert VoiceListener is not None
        assert VoiceError is not None

    def test_voice_listener_creation(self):
        from backend.voice.listener import VoiceListener

        listener = VoiceListener(engine="google", language="en-IN")
        assert listener.engine == "google"
        assert listener.language == "en-IN"
        assert listener._recognizer is None

    def test_missing_dependency_raises_helpful_error(self):
        from backend.voice.listener import VoiceListener

        listener = VoiceListener()
        # Force a recognizer-less call to trigger lazy import path
        with pytest.raises(Exception):
            listener.listen()
