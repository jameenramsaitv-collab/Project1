"""
HR Support FastMCP Server

Exposes HR policy search, policy retrieval, and real/simulated email dispatch tools
via the Model Context Protocol (MCP) using standard I/O (stdio).
"""

import os
import re
import sys
import smtplib
from email.message import EmailMessage
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

# Load .env variables
load_dotenv()

# Initialize FastMCP Server
mcp = FastMCP("HRSupportServer")

BASE_DIR = Path(__file__).parent.resolve()
POLICIES_DIR = BASE_DIR / "policies"
SENT_EMAILS_DIR = BASE_DIR / "sent_emails"


def _read_all_policies() -> dict[str, str]:
    """Read all markdown files in policies directory."""
    policies = {}
    if not POLICIES_DIR.exists():
        return policies

    for file in POLICIES_DIR.glob("*.md"):
        policies[file.stem] = file.read_text(encoding="utf-8")
    return policies


@mcp.tool()
def list_available_policies() -> list[str]:
    """
    List all available HR policy topics currently registered in the system.
    Returns a list of policy topic identifiers.
    """
    policies = _read_all_policies()
    return sorted(list(policies.keys()))


@mcp.tool()
def get_policy_document(policy_name: str) -> str:
    """
    Retrieve the full raw markdown content of a specific policy document.
    
    Args:
        policy_name: The name/slug of the policy (e.g. 'leave_policy', 'remote_work_policy', 'health_benefits')
    """
    clean_name = policy_name.replace(".md", "").strip().lower()
    target_file = POLICIES_DIR / f"{clean_name}.md"

    if not target_file.exists():
        available = list_available_policies()
        return (
            f"Error: Policy document '{policy_name}' not found. "
            f"Available policies: {', '.join(available)}"
        )

    return target_file.read_text(encoding="utf-8")


@mcp.tool()
def search_policies(query: str) -> str:
    """
    Search across all HR policy documents for sections relevant to a specific query.
    Extracts and ranks matching clauses with citations.

    Args:
        query: The search term or question (e.g. 'maternity leave', 'gym reimbursement', 'work from home equipment')
    """
    policies = _read_all_policies()
    if not policies:
        return "No policy documents found in the system."

    keywords = [k.lower() for k in re.findall(r"\w+", query) if len(k) > 2]
    if not keywords:
        return "Search query was too short or contained no valid keywords."

    matched_sections = []

    for doc_name, content in policies.items():
        # Split document by markdown section headers (##)
        raw_sections = re.split(r"\n(?=##\s)", content)
        for section in raw_sections:
            section_lower = section.lower()
            score = sum(section_lower.count(kw) for kw in keywords)
            if score > 0:
                first_line = section.strip().split("\n")[0]
                matched_sections.append({
                    "doc_name": doc_name,
                    "header": first_line.replace("#", "").strip(),
                    "content": section.strip(),
                    "score": score
                })

    if not matched_sections:
        return (
            f"No specific policy sections matched '{query}'. "
            f"You can review available policies using list_available_policies."
        )

    matched_sections.sort(key=lambda x: x["score"], reverse=True)

    results = []
    for item in matched_sections[:3]:
        results.append(
            f"--- Citation: [{item['doc_name']}.md] Section: {item['header']} ---\n"
            f"{item['content']}\n"
        )

    return "\n".join(results)


@mcp.tool()
def send_policy_email(recipient_email: str, subject: str, body: str) -> str:
    """
    Send an email containing policy information or summaries to an employee.
    If SMTP credentials (SMTP_HOST, SMTP_USER, SMTP_PASSWORD) are set in .env,
    an actual email is delivered to the recipient.
    In all cases, a copy is archived to sent_emails/ for audit and web preview.

    Args:
        recipient_email: The destination email address of the employee (e.g. 'john.doe@company.com')
        subject: The subject line for the email
        body: The formatted content/body of the email (supports markdown/plain text)
    """
    recipient_clean = recipient_email.strip()
    email_regex = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    if not re.match(email_regex, recipient_clean):
        return f"Error: '{recipient_email}' is not a valid email address."

    # 1. Archive email locally
    SENT_EMAILS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_recipient = re.sub(r"[^\w\.-]", "_", recipient_clean)
    filename = f"email_{timestamp_str}_{safe_recipient}.txt"
    filepath = SENT_EMAILS_DIR / filename

    from_address = os.getenv("SMTP_FROM_EMAIL") or os.getenv("SMTP_USER") or "hr-support@acmecorp.internal"

    archive_content = (
        f"====================================================\n"
        f"TIMESTAMP: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"TO:        {recipient_clean}\n"
        f"FROM:      {from_address}\n"
        f"SUBJECT:   {subject}\n"
        f"====================================================\n\n"
        f"{body}\n\n"
        f"--\n"
        f"Acme Corp HR Support Agent\n"
        f"Internal Portal: https://hr.acmecorp.internal\n"
    )
    filepath.write_text(archive_content, encoding="utf-8")

    # 2. Check for real SMTP configuration
    smtp_host = os.getenv("SMTP_HOST", "").strip()
    smtp_port_str = os.getenv("SMTP_PORT", "587").strip()
    smtp_user = os.getenv("SMTP_USER", "").strip()
    smtp_password = os.getenv("SMTP_PASSWORD", "").strip()

    if smtp_host and smtp_user and smtp_password:
        try:
            port = int(smtp_port_str)
            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = from_address
            msg["To"] = recipient_clean
            msg.set_content(
                f"{body}\n\n--\nAcme Corp HR Support Agent\nInternal Portal: https://hr.acmecorp.internal"
            )

            # Connect via SSL or STARTTLS
            if port == 465:
                with smtplib.SMTP_SSL(smtp_host, port, timeout=15) as server:
                    server.login(smtp_user, smtp_password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(smtp_host, port, timeout=15) as server:
                    server.starttls()
                    server.login(smtp_user, smtp_password)
                    server.send_message(msg)

            sys.stderr.write(f"\n[MCP Server] Real email dispatched via SMTP to {recipient_clean}.\n")
            sys.stderr.flush()

            return (
                f"Real email successfully dispatched to '{recipient_clean}' via SMTP ({smtp_host})! "
                f"Subject: '{subject}'. "
                f"(Archived in sent_emails/{filename})"
            )

        except Exception as e:
            sys.stderr.write(f"\n[MCP Server SMTP Error]: {e}\n")
            sys.stderr.flush()
            return (
                f"Email saved locally to sent_emails/{filename} for '{recipient_clean}', "
                f"but live SMTP delivery encountered an error: {e}. "
                f"Please verify your SMTP credentials in .env."
            )

    # If SMTP is not configured, inform the user cleanly
    sys.stderr.write(f"\n[MCP Server] Email archived locally to {filename} (SMTP not configured in .env)\n")
    sys.stderr.flush()
    return (
        f"Email successfully composed and archived to sent_emails/{filename} for '{recipient_clean}'. "
        f"Subject: '{subject}'. "
        f"(Note: To deliver real emails to inboxes, specify SMTP_HOST, SMTP_USER, and SMTP_PASSWORD in .env)"
    )


@mcp.resource("policy://{policy_name}")
def get_policy_resource(policy_name: str) -> str:
    """Read a policy document directly as an MCP resource."""
    return get_policy_document(policy_name)


if __name__ == "__main__":
    mcp.run()
