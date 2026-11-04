"""Smoke tests for Step 1 — confirms the app starts and home page loads."""

import pytest
from httpx import AsyncClient, ASGITransport
from app.main import app


@pytest.mark.asyncio
async def test_home_page_returns_200():
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        response = await client.get("/")
    assert response.status_code == 200


@pytest.mark.asyncio
async def test_home_page_contains_h1():
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        response = await client.get("/")
    assert "<h1" in response.text


@pytest.mark.asyncio
async def test_home_page_contains_skip_link():
    """Skip link is required for keyboard users to bypass navigation."""
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        response = await client.get("/")
    assert "skip-link" in response.text
    assert "#main-content" in response.text


@pytest.mark.asyncio
async def test_home_page_has_lang_attribute():
    """lang attribute is required for screen readers to use the correct language."""
    async with AsyncClient(transport=ASGITransport(app=app), base_url="http://test") as client:
        response = await client.get("/")
    assert 'lang="en"' in response.text
