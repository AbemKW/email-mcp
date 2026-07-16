"""COM layer: Outlook session, dedicated STA worker, and folder helpers.

Everything COM-touching lives here. The ops layer receives a live
:class:`~email_mcp.outlook.session.OutlookSession` and never constructs COM itself,
which keeps ops thread-agnostic and testable against a fake session.
"""

from email_mcp.outlook.session import OutlookSession
from email_mcp.outlook.worker import OutlookWorker

__all__ = ["OutlookSession", "OutlookWorker"]
