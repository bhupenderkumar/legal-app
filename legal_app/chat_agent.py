"""
Legal Chat Agent — Conversational AI using Amazon Nova Pro on AWS Bedrock.
Enforces data-gathering before document generation, tracks tool usage,
and supports attorney profile integration.
"""

import os
import json
import re
import uuid

import boto3

from document_gen import AttorneyProfile

bedrock = boto3.client("bedrock-runtime", region_name=os.environ.get("AWS_REGION", "us-east-1"))
MODEL = os.environ.get("LEXAI_MODEL", "amazon.nova-pro-v1:0")

SYSTEM_PROMPT = """You are LexAI, an expert US legal assistant built for a law firm. Your job is to help the firm's clients and the attorney with:

1. **Legal advice** — Analyze situations, explain rights, suggest strategies
2. **Document drafting** — Create court-ready legal documents (complaints, demand letters, motions, NDAs, contracts, etc.)
3. **Legal research** — Look up statutes, case law, and legal principles

## Rules
- Be conversational, thorough, and professional. You represent a law firm.
- Always include a disclaimer that you're an AI and this is not legal advice.
- Use the tools available to you when research or document generation is needed.
- **For document drafting**: gather ALL details through conversation before calling generate_legal_document. You need party names, jurisdiction/state, detailed facts, and relief sought.

## Attorney Profile
The firm's information is shown below. Documents should be prepared on behalf of this attorney/firm. The system adds letterhead and signature blocks automatically — do NOT include them in your generated text, just the body content."""

AGENT_TOOLS = [
    {
        "toolSpec": {
            "name": "lookup_statute",
            "description": "Look up a law, statute, or legal principle by topic. Use to research the law before giving advice or drafting documents.",
            "inputSchema": {
                "json": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "Legal topic or statute to look up (e.g., 'breach of contract UCC', 'non-compete California')"
                        },
                        "jurisdiction": {
                            "type": "string",
                            "description": "Jurisdiction (e.g., 'US Federal', 'California', 'New York')"
                        }
                    },
                    "required": ["query"]
                }
            }
        }
    },
    {
        "toolSpec": {
            "name": "search_case_law",
            "description": "Search for relevant case law precedents on a legal issue. Returns case names, holdings, and relevance.",
            "inputSchema": {
                "json": {
                    "type": "object",
                    "properties": {
                        "legal_issue": {
                            "type": "string",
                            "description": "The specific legal issue to find precedents for (be precise)"
                        },
                        "jurisdiction": {
                            "type": "string",
                            "description": "Jurisdiction to search in"
                        }
                    },
                    "required": ["legal_issue"]
                }
            }
        }
    },
    {
        "toolSpec": {
            "name": "analyze_contract_clause",
            "description": "Analyze a contract clause for enforceability, risks, and recommendations. Use when a user shares contract text.",
            "inputSchema": {
                "json": {
                    "type": "object",
                    "properties": {
                        "clause_text": {
                            "type": "string",
                            "description": "The contract clause text to analyze"
                        },
                        "clause_type": {
                            "type": "string",
                            "enum": ["non-compete", "indemnification", "termination", "liability", "confidentiality", "force_majeure", "other"],
                            "description": "Type of clause"
                        }
                    },
                    "required": ["clause_text"]
                }
            }
        }
    },
    {
        "toolSpec": {
            "name": "generate_legal_document",
            "description": "Generate a complete legal document ready for download. Use ONLY after gathering ALL required info: party names, jurisdiction/state, detailed facts, and relief sought. NEVER use placeholder data.",
            "inputSchema": {
                "json": {
                    "type": "object",
                    "properties": {
                        "document_type": {
                            "type": "string",
                            "enum": ["demand_letter", "complaint", "answer", "motion", "affidavit", "cease_desist", "nda", "contract"]
                        },
                        "plaintiff_name": {"type": "string"},
                        "plaintiff_address": {"type": "string"},
                        "defendant_name": {"type": "string"},
                        "defendant_address": {"type": "string"},
                        "state": {"type": "string"},
                        "court_name": {"type": "string"},
                        "case_number": {"type": "string"},
                        "claim_type": {"type": "string"},
                        "amount": {"type": "string"},
                        "facts": {"type": "string", "description": "Detailed factual background"},
                        "relief_sought": {"type": "string"},
                        "deadline_days": {"type": "string"},
                        "additional_terms": {"type": "string"}
                    },
                    "required": ["document_type", "facts"]
                }
            }
        }
    }
]


def lookup_statute(query, jurisdiction="US Federal"):
    db = {
        "breach of contract": {
            "title": "Uniform Commercial Code (UCC) § 2-711 to 2-717",
            "jurisdiction": "US Federal / All 50 states",
            "summary": "Governs buyer's remedies for seller's breach in sale of goods.",
            "key_elements": ["Valid enforceable contract", "Plaintiff's performance or valid excuse", "Material breach by defendant", "Damages proximately caused by breach"],
            "statute_of_limitations": "4-6 years for written contracts (varies by state)"
        },
        "non-compete": {
            "title": "Non-Compete Agreement Enforceability",
            "jurisdiction": jurisdiction,
            "summary": "Enforceability varies by state. Most apply reasonableness test.",
            "key_elements": ["Adequate consideration", "Reasonable time (6mo-2yr typical)", "Reasonable geographic scope", "Reasonable activity restriction", "Legitimate business interest"],
            "state_variations": {
                "California": "Bus. & Prof. Code § 16600 — VOID",
                "New York": "Strict reasonableness; courts may blue-pencil",
                "Texas": "Bus. & Com. Code § 15.50 — Enforceable if reasonable",
                "Florida": "Fla. Stat. § 542.335 — Presumed reasonable if <= 2 years"
            }
        },
        "tenant rights": {
            "title": "Fair Housing Act (42 U.S.C. §§ 3601-3619) & State Landlord-Tenant Laws",
            "jurisdiction": jurisdiction,
            "summary": "Comprehensive tenant protections: implied warranty of habitability, prohibition on self-help eviction, security deposit rules, privacy rights, anti-retaliation, anti-discrimination.",
            "key_elements": ["Implied warranty of habitability", "Prohibition on self-help eviction", "Security deposit limits", "Right to privacy / notice of entry", "Anti-retaliation protections"]
        },
        "intellectual property": {
            "title": "U.S. Intellectual Property Framework",
            "jurisdiction": "US Federal",
            "summary": "IP protection through copyright (17 U.S.C.), trademark (Lanham Act), patent (35 U.S.C.), and trade secrets (DTSA).",
            "key_elements": ["Copyright: automatic on creation", "Trademark: use in commerce; USPTO registration", "Patent: 20-year term", "Trade Secrets: reasonable secrecy measures"]
        },
        "employment": {
            "title": "Federal & State Employment Law",
            "jurisdiction": jurisdiction,
            "summary": "Employment protections under Title VII, FLSA, FMLA, ADA, ADEA, OSHA.",
            "key_elements": ["At-will employment (default in most states)", "Anti-discrimination", "Wage and hour (FLSA)", "Family leave (FMLA)", "Wrongful termination exceptions"]
        }
    }
    for key, value in db.items():
        if key in query.lower():
            return json.dumps(value, indent=2)
    return json.dumps({
        "title": f"Search: {query}",
        "jurisdiction": jurisdiction,
        "summary": "General legal principles apply. Jurisdiction-specific research recommended.",
        "note": "For production, connect to Westlaw or LexisNexis."
    }, indent=2)


def search_case_law(legal_issue, jurisdiction="US Federal"):
    db = {
        "breach of contract": [
            {"case": "Hadley v. Baxendale, 9 Exch. 341 (1854)", "holding": "Damages limited to those arising naturally or reasonably foreseeable."},
            {"case": "Jacob & Youngs v. Kent, 230 N.Y. 239 (1921)", "holding": "Substantial performance excuses minor deviations."},
        ],
        "non-compete": [
            {"case": "Edwards v. Arthur Andersen LLP, 44 Cal.4th 937 (2008)", "holding": "Bus. & Prof. Code § 16600 invalidates employee non-competes."},
            {"case": "BDO Seidman v. Hirshberg, 93 N.Y.2d 382 (1999)", "holding": "Non-competes enforceable if necessary, not burdensome, not harmful."}
        ],
        "eviction": [
            {"case": "Javins v. First National Realty Corp., 428 F.2d 1071 (D.C. Cir. 1970)", "holding": "Implied warranty of habitability in leases."},
            {"case": "Robinson v. Diamond Housing Corp., 463 F.2d 853 (D.C. Cir. 1972)", "holding": "Retaliatory eviction prohibited."}
        ],
        "negligence": [
            {"case": "Palsgraf v. Long Island Railroad, 248 N.Y. 339 (1928)", "holding": "Duty owed only to foreseeable plaintiffs in zone of danger."}
        ]
    }
    for key, cases in db.items():
        if key in legal_issue.lower():
            return json.dumps({"issue": legal_issue, "cases": cases}, indent=2)
    return json.dumps({"issue": legal_issue, "cases": [], "note": "No matches. Use Westlaw/LexisNexis for production."}, indent=2)


def analyze_contract_clause(clause_text, clause_type=None):
    analysis = {"clause_type": clause_type or "general", "risk_assessment": {}, "red_flags": [], "recommendations": []}
    text_lower = clause_text.lower()
    time_match = re.search(r'(\d+)\s*(year|month|day)', text_lower)
    if time_match:
        num, unit = int(time_match.group(1)), time_match.group(2)
        if unit == "year" and num > 2:
            analysis["red_flags"].append(f"Duration of {num} years exceeds typical threshold")
            analysis["risk_assessment"]["duration"] = "HIGH"
        else:
            analysis["risk_assessment"]["duration"] = "LOW"
    geo_match = re.search(r'(\d+)\s*mile', text_lower)
    if geo_match:
        miles = int(geo_match.group(1))
        if miles > 100:
            analysis["red_flags"].append(f"Radius of {miles} miles is excessively broad")
            analysis["risk_assessment"]["geographic_scope"] = "HIGH"
    vague_terms = ["any competitor", "any business", "any capacity", "in any way"]
    found = [t for t in vague_terms if t in text_lower]
    if found:
        analysis["red_flags"].append(f"Overbroad terms: {', '.join(found)}")
        analysis["risk_assessment"]["specificity"] = "HIGH"
    high_count = sum(1 for v in analysis["risk_assessment"].values() if v == "HIGH")
    analysis["overall_risk"] = "HIGH" if high_count >= 2 else "MEDIUM-HIGH" if high_count == 1 else "MEDIUM"
    analysis["recommendations"] = ["Negotiate narrower restrictions", "Add severability clause", "Specify governing law", "Include carve-out for general skills"]
    return json.dumps(analysis, indent=2)


def _is_empty_placeholder(text: str) -> bool:
    placeholders = ["your name", "unknown", "placeholder", "n/a", "not provided", "tbd", "to be determined"]
    t = text.strip().lower()
    if len(t) < 3:
        return True
    for p in placeholders:
        if p == t:
            return True
    return False


def generate_legal_document(**kwargs):
    doc_type = kwargs.get("document_type", "demand_letter")
    parts = [f"Generate a complete, professional {doc_type.replace('_', ' ').title()} with the following details:\n"]
    for key, value in kwargs.items():
        if value and key != "document_type":
            label = key.replace("_", " ").title()
            parts.append(f"{label}: {value}")
    user_prompt = "\n".join(parts)

    DOC_PROMPTS = {
        "demand_letter": """You are an aggressive senior litigation attorney drafting a FORMAL DEMAND LETTER.

REQUIRED STRUCTURE:
1. SENDER BLOCK: Client name and address
2. RECIPIENT BLOCK: Opposing party name and address
3. RE LINE: Subject line
4. SALUTATION
5. OPENING: State you represent the client, cite relevant statute
6. DETAILED FACTS: Chronological narrative with dates and amounts
7. LEGAL VIOLATIONS: Cite 3-4 specific legal theories
8. DAMAGES: Exact dollar amount, treble damages, fees, interest
9. REGULATORY THREATS
10. FIRM DEADLINE
11. LITIGATION WARNING
12. CLOSING
TONE: Professional but firm. Output ONLY the letter text. No markdown. No brackets.""",

        "complaint": """You are a senior litigation attorney drafting a CIVIL COMPLAINT following FRCP Rule 8.

STRUCTURE:
1. CAPTION: Court name, parties, case number placeholder
2. NATURE OF ACTION
3. PARTIES: Full identification
4. JURISDICTION AND VENUE: Cite statutes
5. FACTUAL ALLEGATIONS: Numbered paragraphs
6. CAUSES OF ACTION: Each as separate COUNT with statutory authority
7. PRAYER FOR RELIEF
8. JURY DEMAND
Output ONLY the complaint text. No markdown.""",

        "answer": """Draft an Answer. ADMIT true facts, DENY false ones, state INSUFFICIENT KNOWLEDGE. Raise ALL affirmative defenses. Output ONLY the answer text.""",

        "motion": """Draft a court motion with memorandum of points and authorities. Include: caption, introduction, facts, legal standard, argument with case citations, conclusion, proposed order. Output ONLY the motion text.""",

        "affidavit": """Draft a sworn declaration under 28 U.S.C. § 1746. Numbered paragraphs, personal knowledge basis. Output ONLY the affidavit text.""",

        "cease_desist": """Draft a CEASE AND DESIST letter. Identify rights violated, describe infringement, cite legal authority, demand cessation, set deadline, warn of injunction. Output ONLY the letter text.""",

        "nda": """Draft a comprehensive Non-Disclosure Agreement. Include: parties, recitals, Confidential Information definition, obligations, exclusions, term, remedies, governing law, signature blocks. Output ONLY the agreement text.""",

        "contract": """Draft a comprehensive service/goods agreement. Include: parties, scope, payment terms, warranties, indemnification, limitation of liability, termination, force majeure, dispute resolution, governing law, signature blocks. Output ONLY the contract text.""",
    }

    system = DOC_PROMPTS.get(doc_type, DOC_PROMPTS["demand_letter"])
    system += "\n\nIMPORTANT: Do NOT include letterhead, attorney signature block, or firm info. The system adds those automatically."

    response = bedrock.converse(
        modelId=MODEL,
        messages=[{"role": "user", "content": [{"text": user_prompt}]}],
        system=[{"text": system}],
        inferenceConfig={"maxTokens": 4096, "temperature": 0.3},
    )
    doc_text = "".join(b["text"] for b in response["output"]["message"]["content"] if "text" in b)

    doc_id = str(uuid.uuid4())[:8]
    return json.dumps({
        "status": "document_ready",
        "document_id": doc_id,
        "document_type": doc_type,
        "content": doc_text,
        "message": f"Your {doc_type.replace('_', ' ')} has been generated."
    })


TOOL_DISPATCH = {
    "lookup_statute": lookup_statute,
    "search_case_law": search_case_law,
    "analyze_contract_clause": analyze_contract_clause,
    "generate_legal_document": generate_legal_document,
}

MAX_TOKENS = 4096


class ChatSession:
    def __init__(self):
        self.messages = []
        self.documents = {}
        self._turn_count = 0
        self.profile = AttorneyProfile()

    def _system_prompt(self) -> str:
        if self.profile.is_empty:
            return SYSTEM_PROMPT + "\n\nNOTE: No attorney profile has been set up yet."
        lines = ["\n\n## Current Firm Profile"]
        if self.profile.firm_name:
            lines.append(f"Firm: {self.profile.firm_name}")
        if self.profile.attorney_name:
            lines.append(f"Attorney: {self.profile.attorney_name}")
        if self.profile.address:
            lines.append(f"Address: {self.profile.address}")
        if self.profile.phone:
            lines.append(f"Phone: {self.profile.phone}")
        if self.profile.email:
            lines.append(f"Email: {self.profile.email}")
        if self.profile.bar_number:
            lines.append(f"Bar No.: {self.profile.bar_number}")
        lines.append("\nDocuments are prepared on behalf of this firm. The system adds letterhead and signature blocks automatically. Do NOT include them in your generated text.")
        return SYSTEM_PROMPT + "\n".join(lines)

    def _can_use_tool(self, fn_name: str, fn_args: dict) -> bool:
        if fn_name == "generate_legal_document":
            if self._turn_count < 2:
                return False
            facts = fn_args.get("facts", "")
            if len(facts.strip()) < 20 or _is_empty_placeholder(facts):
                return False
        return True

    @staticmethod
    def _summarize(val, max_len=200):
        s = json.dumps(val, indent=2) if isinstance(val, dict) else str(val)
        if len(s) > max_len:
            s = s[:max_len] + "..."
        return s

    def chat(self, user_message: str) -> dict:
        self.messages.append({
            "role": "user",
            "content": [{"text": user_message}]
        })

        self._turn_count += 1
        document = None
        tools_used = []

        for _ in range(10):
            response = bedrock.converse(
                modelId=MODEL,
                messages=self.messages,
                system=[{"text": self._system_prompt()}],
                inferenceConfig={"maxTokens": MAX_TOKENS},
                toolConfig={"tools": AGENT_TOOLS},
            )

            output_msg = response["output"]["message"]
            stop_reason = response.get("stopReason", "")

            if stop_reason == "tool_use":
                self.messages.append(output_msg)
                tool_results = []

                for block in output_msg["content"]:
                    if "toolUse" not in block:
                        continue
                    tu = block["toolUse"]
                    fn_name = tu["name"]
                    fn_args = tu["input"]
                    tool_use_id = tu["toolUseId"]

                    if not self._can_use_tool(fn_name, fn_args):
                        result = json.dumps({
                            "status": "blocked",
                            "reason": "I need more information before generating the document."
                        })
                        tools_used.append({
                            "name": fn_name,
                            "args": self._summarize(fn_args),
                            "status": "blocked",
                            "result": "More info required",
                        })
                    else:
                        fn = TOOL_DISPATCH.get(fn_name)
                        result = fn(**fn_args) if fn else json.dumps({"error": f"Unknown tool: {fn_name}"})
                        tools_used.append({
                            "name": fn_name,
                            "args": self._summarize(fn_args),
                            "status": "done",
                            "result": self._summarize(result, 300),
                        })

                    if fn_name == "generate_legal_document":
                        try:
                            doc_data = json.loads(result)
                            if doc_data.get("status") == "document_ready":
                                doc_id = doc_data["document_id"]
                                self.documents[doc_id] = {
                                    "type": doc_data["document_type"],
                                    "content": doc_data["content"],
                                }
                                document = {"id": doc_id, "type": doc_data["document_type"]}
                        except json.JSONDecodeError:
                            pass

                    tool_results.append({
                        "toolResult": {
                            "toolUseId": tool_use_id,
                            "content": [{"text": result}],
                        }
                    })

                self.messages.append({"role": "user", "content": tool_results})
            else:
                reply = ""
                for block in output_msg["content"]:
                    if "text" in block:
                        reply += block["text"]
                reply = re.sub(r'<thinking>.*?</thinking>\s*', '', reply, flags=re.DOTALL)
                self.messages.append(output_msg)
                return {"reply": reply, "document": document, "tools_used": tools_used}

        return {"reply": "I'm still working on this. Could you try again?", "document": None, "tools_used": tools_used}
