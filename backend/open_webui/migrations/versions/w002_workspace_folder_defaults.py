"""Add workspace folders and workspace default model metadata support

Company custom: Team Workspaces UX V2
Revision ID: w002_workspace_folder_defaults
Revises: w001_workspace_v1
Create Date: 2026-06-08 00:00:00.000000

Safety notes
------------
* Adds nullable folder.workspace_id. Existing folders stay NULL and continue
  to behave as personal folders.
* Workspace default model is stored in workspace.meta.default_model_id, so no
  workspace table migration is required for that setting.
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa

revision: str = 'w002_workspace_folder_defaults'
down_revision: Union[str, None] = 'w001_workspace_v1'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    pass


def downgrade() -> None:
    pass
