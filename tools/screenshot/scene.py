#!/usr/bin/env python3
"""What the README screenshot shows: an invented sample project with phases,
steps, actions, and a contact.

None of this is real data. The version shown in the image always comes from
simple_project_manager.py, never from here, so this fixture holds no version
number.
"""

STATUSES = ["Not started", "In progress", "Done", "Blocked"]
PRIORITIES = ["Low", "Normal", "High", "Critical"]


def _item(item_id, item_type, name, **kw):
    base = dict(id=item_id, type=item_type, name=name, owner="", coowner="",
                deadline="", status="" if item_type in ("Phase", "Contact")
                else "Not started", priority="Normal", tags=[], notes="")
    base.update(kw)
    return base


ITEMS = [
    _item("id0", "Phase", "Foundation", tags=["setup"]),
    _item("id1", "Step", "Provision hardware", owner="jde", status="Done",
          priority="High", deadline="2026-07-10", tags=["infra"]),
    _item("id2", "Action", "Rack and cable the node", status="Done"),
    _item("id3", "Action", "Label ports", status="Done"),
    _item("id4", "Step", "Base OS and hardening", owner="jde", coowner="sam",
          status="In progress", priority="Critical", deadline="2026-08-01",
          tags=["security"]),
    _item("id5", "Action", "Apply CIS baseline", status="In progress"),
    _item("id6", "Action", "Enable disk encryption", status="Not started"),
    _item("id7", "Action", "Configure firewall", status="Blocked"),
    _item("id8", "Phase", "Rollout", tags=["release"]),
    _item("id9", "Step", "Pilot deployment", owner="sam",
          status="In progress", priority="Normal", deadline="2026-08-20"),
    _item("id10", "Action", "Deploy to staging", status="Done"),
    _item("id11", "Action", "Smoke test", status="In progress"),
    _item("id12", "Contact", "On-call ops",
          notes="Platform team · ops@corp.example · +1-555-0142"),
    _item("id13", "Step", "Documentation", status="Not started",
          priority="Low"),
    _item("id14", "Action", "Write runbook", status="Not started"),
]
