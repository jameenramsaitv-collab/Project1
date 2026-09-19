How Policy Lookup Works (Step by Step)
Step 1 — The LLM decides which tool to call
When you type a query like "How many parental leave days do I get?", it goes to the Google Antigravity Agent (Gemini). The model reads the system instructions which describe all 4 available MCP tools. Gemini uses its reasoning to decide:

"This is a policy question. I should call search_policies(query='parental leave days')."

Step 2 — Keyword extraction from the query (inside hr_mcp_server.py)
In 
search_policies()
, the query is broken into individual keywords — words longer than 2 characters:

python


keywords = [k.lower() for k in re.findall(r"\w+", query) if len(k) > 2]
# "parental leave days" → ["parental", "leave", "days"]
Step 3 — All 3 policy files are read and split into sections
All policies/*.md files are loaded from disk. Each file is then split into individual sections by ## markdown heading:

python


raw_sections = re.split(r"\n(?=##\s)", content)
So leave_policy.md becomes separate chunks like:

## 1. Overview
## 2. Paid Time Off (PTO) / Annual Leave
## 3. Sick & Medical Leave
## 4. Parental Leave
## 5. Bereavement Leave
etc.
Step 4 — Keyword frequency scoring (relevance ranking)
Each section chunk (not the whole file) is scored by counting how many times the keywords appear in it:

python


score = sum(section_lower.count(kw) for kw in keywords)
For the query "parental leave days", the section ## 4. Parental Leave would score much higher than ## 1. Overview.

Step 5 — Only the top 3 matching sections are returned
python


matched_sections.sort(key=lambda x: x["score"], reverse=True)
# Returns top 3 sections with their filename citation
So the LLM does NOT receive the full content of all 3 policy files. It receives only the top 3 most relevant section excerpts, each tagged with its source filename and section header, e.g.:



--- Citation: [leave_policy.md] Section: 4. Parental Leave ---
## 4. Parental Leave
Acme Corp supports growing families...
- Primary Caregiver Leave: 16 weeks of 100% paid leave...
Step 6 — The LLM synthesizes the answer
The Gemini model receives this partial, pre-filtered text as the tool result and uses it to compose a natural, well-cited response. It does not write the raw tool output back to you — it reads the excerpts and generates a helpful human-readable summary in its own words.

Summary Table
Question	Answer
Does it search all policy files?	✅ Yes, all .md files in policies/ are scanned.
Does it send full file content to LLM?	❌ No — only top 3 relevant section chunks.
How are sections ranked?	By keyword frequency count (simple lexical scoring).
Does the LLM make the final decision?	✅ Yes — it synthesizes, formats, and cites the returned chunks.
One Limitation Worth Noting
The current search uses simple keyword frequency matching (lexical). This works well for direct queries but can miss semantically similar terms (e.g., searching "maternity" won't match the policy that uses the word "parental"). For a production system, you could replace this with a proper vector embedding / semantic search (e.g., using sentence-transformers + FAISS or ChromaDB).