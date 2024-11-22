import pytest
from app import create_app

@pytest.fixture
def client():
    app = create_app()
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client

def test_index_route(client):
    """Test the default route."""
    response = client.get('/')
    assert response.status_code == 200
    assert b"Welcome to the Statement Generator API" in response.data
