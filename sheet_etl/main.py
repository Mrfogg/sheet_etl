from autogen_agentchat.agents import AssistantAgent
from autogen_agentchat.conditions import TextMentionTermination
from autogen_agentchat.teams import RoundRobinGroupChat
from autogen_agentchat.ui import Console
from autogen_ext.models.openai import OpenAIChatCompletionClient
import asyncio
model_client = OpenAIChatCompletionClient(
    model="doubao-seed-2-0-mini-260215",
    api_key="28f93d70-ea9c-47f8-8cce-5eae9e8cf3bc",
    base_url="https://ark.cn-beijing.volces.com/api/v3",
    model_info={
        "vision": False,
        "function_calling": True,
        "json_output": True,
        "family": "unknown",
        "structured_output": True
    },
)

planner_agent = AssistantAgent(
    "planner_agent",
    model_client=model_client,
    description="A helpful assistant that can plan trips.",
    system_message="You are a helpful assistant that can suggest a travel plan for a user based on their request.",
)

async def main() -> None:
    termination = TextMentionTermination("TERMINATE")
    group_chat = RoundRobinGroupChat(
        [planner_agent], termination_condition=termination
    )
    await Console(group_chat.run_stream(task="Plan a 3 day trip to Nepal."))

    await model_client.close()

asyncio.run(main())