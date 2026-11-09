"""Unreadable, undeletable recovery copy for the recovery-locked-unreadable scenario."""
from fixture_recovery_locked import hold_recovery

if __name__ == "__main__":
    hold_recovery("recovery-locked-unreadable", unreadable=True)
