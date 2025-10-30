# Copyright (c) Meta Platforms, Inc. and affiliates.
# All rights reserved.
#
# This source code is licensed under the BSD-style license found in the
# LICENSE file in the root directory of this source tree.

"""
FastAPI application for the LibreOffice Spreadsheet Environment.

This module creates an HTTP server that exposes the LibreOfficeEnvironment
over HTTP endpoints, making it compatible with HTTPEnvClient.

Usage:
    # Development (with auto-reload):
    uvicorn envs.libreoffice_env.server.app:app --reload --host 0.0.0.0 --port 8000

    # Production:
    uvicorn envs.libreoffice_env.server.app:app --host 0.0.0.0 --port 8000 --workers 4

    # Or run directly:
    python -m envs.libreoffice_env.server.app
"""

from core.env_server.http_server import create_app

from ..models import LibreOfficeAction, LibreOfficeObservation
from .libreoffice_environment import LibreOfficeEnvironment
import ConfigParser

configParser = ConfigParser.RawConfigParser()   
base_config_file_path,goal_config_file_path = r'BASE_FILE',r'GOAL_FILE'
base_file, goal_file=configParser.read(base_config_file_path),configParser.read(goal_config_file_path)
# Create the environment instance
env = LibreOfficeEnvironment(base_excel=base_file, goal_excel=goal_file)

# Create the app with web interface and README integration
app = create_app(env, LibreOfficeAction, LibreOfficeObservation, env_name="libreoffice_env")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
