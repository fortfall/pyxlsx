import logging
from pytest import fixture

logger = logging.getLogger(__name__)

def test_path(data_path):
    logger.debug(data_path.resolve())
    assert not data_path.is_absolute()
