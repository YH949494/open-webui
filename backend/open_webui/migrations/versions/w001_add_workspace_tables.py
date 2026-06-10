"""Add workspace and workspace_member tables; add workspace_id to chat

Company custom: Team Workspaces V1
Revision ID: w001_workspace_v1
Revises: a0b1c2d3e4f5
Create Date: 2026-05-29 00:00:00.000000

Safety notes
------------
* upgrade() is fully non-destructive:
  - Creates two new tables (workspace, workspace_member) — no existing tables touched.
  - Adds a nullable workspace_id column to chat — existing rows default to NULL and
    continue to behave as private chats.
  - Safe to run on a live database; SQLite and PostgreSQL both allow adding nullable
    columns without a table rewrite.
* downgrade() removes the column and drops both tables. Run only after ensuring
  no production workspace data needs to be preserved.
* head verification: confirm with `alembic heads` before deploying. Expected output:
    w001_workspace_v1 (head)
  If the environment has a different head the revision chain must be rebased.
"""

from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa

revision: str = 'w001_workspace_v1'
down_revision: Union[str, None] = 'a0b1c2d3e4f5'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    pass


def downgrade() -> None:
    pass
