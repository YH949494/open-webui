"""Merge obsolete workspace migration branch with upstream head

Revision ID: merge_workspace_obsolete_heads
Revises: 461111b60977, w002_workspace_folder_defaults
Create Date: 2026-06-10
"""

from typing import Sequence, Union

revision: str = "merge_workspace_obsolete_heads"
down_revision: Union[str, tuple[str, str], None] = (
    "461111b60977",
    "w002_workspace_folder_defaults",
)
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    pass


def downgrade() -> None:
    pass
