"""Merge obsolete workspace migration branch with upstream head

Revision ID: merge_workspace_obsolete_heads
Revises: 461111b60977, w002_workspace_folder_defaults
Create Date: 2026-06-10
"""

revision = 'merge_workspace_obsolete_heads'
down_revision = ('461111b60977', 'w002_workspace_folder_defaults')
branch_labels = None
depends_on = None


def upgrade() -> None:
    pass


def downgrade() -> None:
    pass
