"""NTS Config file"""
from ntsparser import buildhash

__all__ = ['settings']


class Settings:
    """Settings repository"""

    prog_name = 'ntsparser'
    version = '1.17-{}'.format(buildhash.commit_short_hash)
    title = 'NTS Parser {}'.format(version)


settings = Settings()
