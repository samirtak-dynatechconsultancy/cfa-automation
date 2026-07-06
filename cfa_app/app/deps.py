"""Singleton wiring for settings, Graph client, settings store and run manager."""

from __future__ import annotations

from functools import lru_cache

from .config import Settings, get_settings
from .graph.auth import TokenProvider
from .graph.client import GraphClient
from .services.run_log import RunLogStore
from .services.runs import RunManager
from .services.settings_store import SettingsStore


@lru_cache
def get_token_provider() -> TokenProvider:
    return TokenProvider(get_settings())


@lru_cache
def get_graph_client() -> GraphClient:
    return GraphClient(get_settings(), get_token_provider())


@lru_cache
def get_settings_store() -> SettingsStore:
    return SettingsStore(get_settings(), get_graph_client())


@lru_cache
def get_run_log_store() -> RunLogStore:
    return RunLogStore(get_settings(), get_graph_client())


@lru_cache
def get_run_manager() -> RunManager:
    return RunManager(get_graph_client(), get_settings_store(), get_run_log_store())


def app_settings() -> Settings:
    return get_settings()
