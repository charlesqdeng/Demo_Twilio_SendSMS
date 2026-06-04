"""Configuration loader for Twilio SMS application."""
import os
from pathlib import Path
from dotenv import load_dotenv


class Config:
    """Load and validate configuration from environment variables."""

    def __init__(self):
        """Initialize configuration by loading .env file."""
        # Load .env file from the same directory
        env_path = Path(__file__).parent / '.env'
        load_dotenv(dotenv_path=env_path)

        # Twilio credentials
        self.twilio_account_sid = os.getenv('TWILIO_ACCOUNT_SID')
        self.twilio_auth_token = os.getenv('TWILIO_AUTH_TOKEN')
        self.from_number = os.getenv('FROM_NUMBER')

        # File paths
        self.template_file = os.getenv('TEMPLATE_FILE', 'message_template.txt')
        self.recipients_file = os.getenv('RECIPIENTS_FILE', 'Sample Data.xlsx')

        self._validate()

    def _validate(self):
        """Validate that all required configuration is present."""
        if not self.twilio_account_sid:
            raise ValueError("TWILIO_ACCOUNT_SID not found in .env file")
        if not self.twilio_auth_token:
            raise ValueError("TWILIO_AUTH_TOKEN not found in .env file")
        if not self.from_number:
            raise ValueError("FROM_NUMBER not found in .env file")

    def get_template(self):
        """Load message template from file."""
        template_path = Path(__file__).parent / self.template_file
        try:
            with open(template_path, 'r') as f:
                return f.read().strip()
        except FileNotFoundError:
            raise FileNotFoundError(f"Template file not found: {template_path}")


def get_config():
    """Get configuration instance."""
    return Config()
