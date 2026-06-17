"""
Legal Chat Agent — Flask App
Chat interface with attorney profile management and document generation.
"""

import os
import logging
from flask import Flask, render_template, request, jsonify, send_file
from chat_agent import ChatSession
from document_gen import (
    AttorneyProfile,
    create_word_document,
    create_pdf_document,
    make_filename,
    create_document_with_integrity,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('legal_app.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.urandom(32)

sessions = {}


def get_session(sid: str) -> ChatSession:
    if sid not in sessions:
        sessions[sid] = ChatSession()
    return sessions[sid]


def get_profile(sid: str) -> AttorneyProfile:
    return get_session(sid).profile


@app.route("/")
def index():
    return render_template("chat.html")


@app.route("/api/profile", methods=["POST"])
def profile():
    data = request.get_json(silent=True) or {}
    sid = data.get("session_id", "default")
    agent = get_session(sid)

    # If firm_name or attorney_name is present, this is a save; otherwise it's a fetch
    if any(data.get(k) for k in ("firm_name", "attorney_name")):
        agent.profile = AttorneyProfile(
            firm_name=data.get("firm_name", ""),
            attorney_name=data.get("attorney_name", ""),
            address=data.get("address", ""),
            phone=data.get("phone", ""),
            email=data.get("email", ""),
            bar_number=data.get("bar_number", ""),
            state=data.get("state", ""),
        )
        return jsonify({"status": "ok", "profile": agent.profile.to_dict()})

    return jsonify({"profile": agent.profile.to_dict(), "is_empty": agent.profile.is_empty})


@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.get_json()
    message = data.get("message", "").strip()
    sid = data.get("session_id", "default")

    if not message:
        return jsonify({"error": "Empty message"}), 400

    agent = get_session(sid)
    result = agent.chat(message)

    response = {"reply": result["reply"], "tools_used": result.get("tools_used", [])}

    if result.get("document"):
        doc = result["document"]
        response["document"] = {
            "id": doc["id"],
            "type": doc["type"],
            "type_label": doc["type"].replace("_", " ").title(),
        }

    return jsonify(response)


@app.route("/api/download/<sid>/<doc_id>/<fmt>")
def download(sid, doc_id, fmt):
    agent = get_session(sid)
    doc = agent.documents.get(doc_id)

    if not doc:
        return jsonify({"error": "Document not found"}), 404

    content = doc["content"]
    doc_type = doc["type"]
    profile = agent.profile
    filename = make_filename(doc_type, "document")

    try:
        if fmt == "docx":
            buffer = create_word_document(content, doc_type, {}, profile)
            return send_file(
                buffer,
                as_attachment=True,
                download_name=f"{filename}.docx",
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        elif fmt == "pdf":
            buffer = create_pdf_document(content, doc_type, {}, profile)
            return send_file(
                buffer,
                as_attachment=True,
                download_name=f"{filename}.pdf",
                mimetype="application/pdf",
            )
        else:
            return jsonify({"error": "Invalid format"}), 400
    except Exception as e:
        logger.error(f"Document download failed: {str(e)}")
        return jsonify({"error": f"Document generation failed: {str(e)}"}), 500


@app.route("/api/reset", methods=["POST"])
def reset():
    data = request.get_json()
    sid = data.get("session_id", "default")
    if sid in sessions:
        del sessions[sid]
    return jsonify({"status": "ok"})


@app.route("/api/session/restore", methods=["POST"])
def restore_session():
    data = request.get_json(silent=True) or {}
    sid = data.get("session_id", "default")
    messages = data.get("messages", [])
    documents = data.get("documents", {})
    profile_data = data.get("profile", {})

    agent = get_session(sid)
    agent.messages = messages

    if profile_data.get("attorney_name") or profile_data.get("firm_name"):
        agent.profile = AttorneyProfile(
            firm_name=profile_data.get("firm_name", ""),
            attorney_name=profile_data.get("attorney_name", ""),
            address=profile_data.get("address", ""),
            phone=profile_data.get("phone", ""),
            email=profile_data.get("email", ""),
            bar_number=profile_data.get("bar_number", ""),
            state=profile_data.get("state", ""),
        )

    for doc_id, doc_info in documents.items():
        agent.documents[doc_id] = doc_info

    return jsonify({"status": "ok", "session_id": sid})


if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  LexAI — Legal Intelligence Platform")
    print("  Powered by Amazon Nova Pro on AWS Bedrock")
    print("  http://localhost:5000")
    print("=" * 60 + "\n")
    app.run(debug=True, port=5000)
