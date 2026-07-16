"""Send-As support — rewrite sender MAPI props so mail is sent AS another address.

Port of ``buildSendAsBlock`` (index.js lines 644-689). Setting only the
PR_SENT_REPRESENTING_* props yields "admin on behalf of contact"; overwriting the
PR_SENDER_* props too is what collapses it to a pure Send As. Exchange validates
the "Send As" permission at submit time — without it the message may bounce, get
rewritten back to "on behalf of", or sit in the Outbox depending on tenant policy.

EXPERIMENTAL: the binary MAPI props (the ``*0102`` proptags: PR_ENTRYID and
PR_SEARCH_KEY) are byte arrays. Whether win32com marshals the value returned by
``PropertyAccessor.GetProperty`` cleanly back into ``SetProperty`` for these binary
proptags is UNVERIFIED against a live Exchange tenant. The string props (``*001F``)
are straightforward; the binary ones are the risk.
"""

from __future__ import annotations

from typing import Any

from email_mcp.errors import SendAsUnresolved

# --- MAPI proptag DASL strings (recipient reads) ---
PR_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
PR_SEARCH_KEY = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"

# --- PR_SENT_REPRESENTING_* (written on the item) ---
PR_SENT_REPRESENTING_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
PR_SENT_REPRESENTING_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"
PR_SENT_REPRESENTING_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x0064001F"
PR_SENT_REPRESENTING_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x00410102"
PR_SENT_REPRESENTING_SEARCH_KEY = "http://schemas.microsoft.com/mapi/proptag/0x003B0102"

# --- PR_SENDER_* (written on the item; overwriting these suppresses "on behalf of") ---
PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
PR_SENDER_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
PR_SENDER_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x0C1E001F"
PR_SENDER_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0C190102"
PR_SENDER_SEARCH_KEY = "http://schemas.microsoft.com/mapi/proptag/0x0C1D0102"


def apply_send_as(session: Any, item: Any, send_as: str) -> None:
    """Rewrite ``item``'s sender MAPI props so it is sent AS ``send_as``.

    No-op if ``send_as`` is falsy. Must be called AFTER ``SendUsingAccount`` is set
    and BEFORE ``.Send()``. Resolves ``send_as`` via the MAPI session; raises
    :class:`SendAsUnresolved` (naming the address) if Exchange cannot resolve it.
    Persists via ``item.Save()``.

    Args:
        session: An :class:`~email_mcp.outlook.session.OutlookSession`.
        item: The mail COM item to rewrite (a fresh ``CreateItem``/reply/forward).
        send_as: The address to send as; falsy means no-op.
    """
    if not send_as:
        return

    recipient = session.app.Session.CreateRecipient(send_as)
    recipient.Resolve()
    if not recipient.Resolved:
        raise SendAsUnresolved(
            f"send_as address could not be resolved by Exchange: {send_as}. "
            "The address must be a recipient Exchange knows about (typically same tenant)."
        )

    ae = recipient.AddressEntry
    rcp_addr_type = str(ae.Type)
    rcp_addr = str(ae.Address)
    rcp_name = str(ae.Name)

    rcp_pa = recipient.PropertyAccessor
    rcp_entry_id = rcp_pa.GetProperty(PR_ENTRYID)
    rcp_search_key = rcp_pa.GetProperty(PR_SEARCH_KEY)

    pa = item.PropertyAccessor
    # PR_SENT_REPRESENTING_{NAME, EMAIL_ADDRESS, ADDRTYPE, ENTRYID, SEARCH_KEY}
    pa.SetProperty(PR_SENT_REPRESENTING_NAME, rcp_name)
    pa.SetProperty(PR_SENT_REPRESENTING_EMAIL_ADDRESS, rcp_addr)
    pa.SetProperty(PR_SENT_REPRESENTING_ADDRTYPE, rcp_addr_type)
    pa.SetProperty(PR_SENT_REPRESENTING_ENTRYID, rcp_entry_id)
    pa.SetProperty(PR_SENT_REPRESENTING_SEARCH_KEY, rcp_search_key)
    # PR_SENDER_{NAME, EMAIL_ADDRESS, ADDRTYPE, ENTRYID, SEARCH_KEY} — overwriting these
    # is what suppresses "on behalf of" and makes it a true Send As.
    pa.SetProperty(PR_SENDER_NAME, rcp_name)
    pa.SetProperty(PR_SENDER_EMAIL_ADDRESS, rcp_addr)
    pa.SetProperty(PR_SENDER_ADDRTYPE, rcp_addr_type)
    pa.SetProperty(PR_SENDER_ENTRYID, rcp_entry_id)
    pa.SetProperty(PR_SENDER_SEARCH_KEY, rcp_search_key)

    item.Save()
