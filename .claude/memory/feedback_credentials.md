---
name: feedback-credentials
description: Never ask the user to paste credentials into the chat; always write to .env files
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 85e53d6d-ff3b-41ca-ab20-c7fe8e7b6e3a
---

When credentials are needed, instruct the user to write them directly to the appropriate env file (e.g., `etl/.env`, `etl/ajera.env`) — never paste into the conversation.

**Why:** Credentials pasted in chat are stored in conversation history. User initially offered to paste them; redirected to file-based approach.

**How to apply:** If a script fails due to missing credentials, open the env file (`open etl/.env`) and ask the user to fill it in, or offer to write a specific key=value to the file directly without echoing the value back in conversation text.
