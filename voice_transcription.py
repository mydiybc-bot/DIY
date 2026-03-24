from __future__ import annotations

import json
import os
import ssl
from uuid import uuid4
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import certifi


AUDIO_TRANSCRIPT_URL = "https://api.openai.com/v1/audio/transcriptions"
DEFAULT_TRANSCRIBE_MODEL = "gpt-4o-mini-transcribe"
DEFAULT_LANGUAGE = "zh"
MAX_AUDIO_BYTES = 25 * 1024 * 1024


class TranscriptionError(RuntimeError):
    pass


def build_multipart_form(
    fields: dict[str, str],
    file_field: str,
    filename: str,
    file_bytes: bytes,
    content_type: str,
) -> tuple[bytes, str]:
    boundary = f"----CodexVoiceBoundary{uuid4().hex}"
    chunks: list[bytes] = []

    for key, value in fields.items():
        chunks.extend(
            [
                f"--{boundary}\r\n".encode("utf-8"),
                f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode("utf-8"),
                str(value).encode("utf-8"),
                b"\r\n",
            ]
        )

    chunks.extend(
        [
            f"--{boundary}\r\n".encode("utf-8"),
            (
                f'Content-Disposition: form-data; name="{file_field}"; filename="{filename}"\r\n'
                f"Content-Type: {content_type}\r\n\r\n"
            ).encode("utf-8"),
            file_bytes,
            b"\r\n",
            f"--{boundary}--\r\n".encode("utf-8"),
        ]
    )

    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def transcribe_audio(audio_bytes: bytes, filename: str, content_type: str) -> str:
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise TranscriptionError("目前尚未設定語音轉文字服務，請先補上 OpenAI API 金鑰。")

    if not audio_bytes:
        raise TranscriptionError("沒有收到語音內容，請再說一次。")

    if len(audio_bytes) > MAX_AUDIO_BYTES:
        raise TranscriptionError("這段錄音太長了，請控制在 25 MB 以內再試一次。")

    model = os.environ.get("OPENAI_TRANSCRIBE_MODEL", DEFAULT_TRANSCRIBE_MODEL).strip() or DEFAULT_TRANSCRIBE_MODEL
    language = os.environ.get("OPENAI_TRANSCRIBE_LANGUAGE", DEFAULT_LANGUAGE).strip()
    fields = {
        "model": model,
        "response_format": "json",
    }
    if language:
        fields["language"] = language

    body, content_type_header = build_multipart_form(
        fields=fields,
        file_field="file",
        filename=filename or "reply.webm",
        file_bytes=audio_bytes,
        content_type=content_type or "application/octet-stream",
    )

    request = Request(AUDIO_TRANSCRIPT_URL, data=body, method="POST")
    request.add_header("Authorization", f"Bearer {api_key}")
    request.add_header("Content-Type", content_type_header)
    request.add_header("Content-Length", str(len(body)))

    try:
        ssl_context = ssl.create_default_context(cafile=certifi.where())
        with urlopen(request, timeout=45, context=ssl_context) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = _read_error_message(exc)
        raise TranscriptionError(f"語音轉文字失敗：{detail}") from exc
    except URLError as exc:
        raise TranscriptionError("目前無法連上語音辨識服務，請稍後再試。") from exc

    transcript = str(payload.get("text", "")).strip()
    if not transcript:
        raise TranscriptionError("這段語音沒有辨識到清楚內容，請再說一次。")
    return transcript


def _read_error_message(exc: HTTPError) -> str:
    try:
        payload = json.loads(exc.read().decode("utf-8"))
        if isinstance(payload, dict):
            error = payload.get("error", {})
            if isinstance(error, dict) and error.get("message"):
                return str(error["message"])
    except Exception:
        pass
    return f"HTTP {exc.code}"
