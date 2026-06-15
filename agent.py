import asyncio
import os
import sys
from dotenv import load_dotenv
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_ollama import ChatOllama
from langgraph.prebuilt import create_react_agent

class ERDAgent:
    def __init__(self):
        load_dotenv()
        self.server_url = os.getenv("MCP_SERVER_URL", "http://localhost:8000/sse")
        self.model_name = os.getenv("OLLAMA_MODEL", "gemma4-agent:latest")
        self.ollama_base_url = os.getenv("OLLAMA_BASE_URL", "http://127.0.0.1:11434")
        self.llm = ChatOllama(model=self.model_name, base_url=self.ollama_base_url)
        self.system_prompt = "You are an expert database design assistant. Use your tools to create ERD tables and connect them. Arrange them nicely using coordinate offsets so they do not overlap."

    async def run(self, prompt: str):
        client = MultiServerMCPClient({
            "microsoft_visio_erd": {
                "transport": "sse",
                "url": self.server_url
            }
        })
        tools = await client.get_tools()
        print("Successfully loaded tools from remote MCP:")
        for tool in tools:
            print(f" - {tool.name}: {tool.description}")
        agent_executor = create_react_agent(
            self.llm,
            tools=tools,
            prompt=self.system_prompt,
        )
        async for chunk in agent_executor.astream({"messages": [("user", prompt)]}):
            for node_name, node_state in chunk.items():
                messages = node_state.get("messages", [])
                for message in messages:
                    message.pretty_print()

    async def start_repl(self):
        while True:
            try:
                prompt = await asyncio.to_thread(input, "ERD Agent> ")
                if not prompt.strip():
                    continue
                if prompt.strip().lower() in ("exit", "quit"):
                    break
                await self.run(prompt)
            except KeyboardInterrupt:
                break

async def main():
    agent = ERDAgent()
    prompt = """Your Prompt"""
    await agent.run(prompt)

if __name__ == "__main__":
    asyncio.run(main())
