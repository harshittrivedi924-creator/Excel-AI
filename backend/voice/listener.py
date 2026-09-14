"""Voice command input for Excel AI Copilot (Phase 9).

Uses speech_recognition (Google Web Speech API by default). The voice layer
only converts speech to text -- all calculation logic stays in the command
engine as required by the roadmap.
"""


class VoiceListener:
    """Convert spoken commands into text using Speech-to-Text."""

    def __init__(self, engine="google", language="en-IN"):
        self.engine = engine
        self.language = language
        self._recognizer = None

    def _get_recognizer(self):
        """Lazily import and build the recognizer so tests can run without
        the optional dependency installed."""
        if self._recognizer is None:
            try:
                import warnings

                with warnings.catch_warnings():
                    warnings.simplefilter("ignore", DeprecationWarning)
                    import speech_recognition as sr
            except ImportError as e:
                raise ImportError(
                    "The 'SpeechRecognition' package is required for voice "
                    "commands. Install it with: pip install SpeechRecognition "
                    "pocketsphinx"
                ) from e
            self._recognizer = sr.Recognizer()
        return self._recognizer

    def listen(self, timeout=5.0, phrase_time_limit=8.0):
        """Listen for a single spoken command and return the recognized text.

        Raises a VoiceError if nothing could be understood or no mic exists.
        """
        recognizer = self._get_recognizer()
        try:
            import speech_recognition as sr

            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=0.3)
                audio = recognizer.listen(
                    source, timeout=timeout, phrase_time_limit=phrase_time_limit
                )
        except Exception as e:
            raise VoiceError(f"Couldn't access the microphone: {e}") from e

        try:
            if self.engine == "google":
                text = recognizer.recognize_google(audio, language=self.language)
            else:
                text = recognizer.recognize_sphinx(audio)
        except sr.UnknownValueError:
            raise VoiceError("I couldn't understand what you said.") from None
        except sr.RequestError as e:
            raise VoiceError(f"Speech service error: {e}") from e

        return text.strip().lower()


class VoiceError(Exception):
    """Raised when voice input fails."""

    def __init__(self, message):
        self.message = message
        super().__init__(message)
