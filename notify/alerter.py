import logging

class Alerter:
    """Dispatches alerts on ETL failures, aborts, or data discrepancies."""

    def __init__(self, logger: logging.Logger):
        self.logger = logger

    def alert_failure(self, run_id: str, step_name: str, error: Exception, context: dict = None) -> None:
        msg = f"[CRITICAL ETL FAILURE] Run ID {run_id} aborted at step '{step_name}': {error}"
        self.logger.critical(msg)
        if context:
            self.logger.critical(f"Failure Context: {context}")
        # Extensible to webhook, Slack, or email notifications
