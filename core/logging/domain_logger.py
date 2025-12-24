"""
domain_logger.py
"""

from core.logging.events import LogEvent,LogLevel

class DomainLogger:
    def __init__(self, logger):
        self._logger = logger
        self.events: list[LogEvent] = []

    def log(self, event: LogEvent):
        # 1. Сохраняем событие для GUI
        self.events.append(event)

        # 2. Логируем в файл / консоль
        text = self._format_event(event)

        if event.level == LogLevel.ERROR:
            self._logger.error(text)
        elif event.level == LogLevel.WARNING:
            self._logger.warning(text)
        else:
            self._logger.info(text)

    def has_errors(self) -> bool:
        return any(e.level == LogLevel.ERROR for e in self.events)

    def _format_event(self, event: LogEvent) -> str:
        parts = [f"[{event.code}]", event.message]

        if event.row is not None:
            parts.append(f"ROW={event.row}")

        if event.raw_value is not None:
            parts.append(f"RAW='{event.raw_value}'")

        if event.parsed_value is not None:
            parts.append(f"PARSED='{event.parsed_value}'")

        return " | ".join(parts)
