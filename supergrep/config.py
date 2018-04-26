"""SuperGrep Config file"""

__all__ = ['settings']


class Settings:
    """Settings repository"""

    prog_name = 'supergrep'
    version = '0.1'
    title = 'SuperGrep version {}'.format(version)


settings = Settings()
