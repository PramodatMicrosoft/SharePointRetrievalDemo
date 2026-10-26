# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from __future__ import annotations

import logging
import re
from os import environ

from dotenv import load_dotenv
from microsoft_agents.activity import Activity, ActivityTypes, Attachment, load_configuration_from_env
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.hosting.aiohttp import CloudAdapter
from microsoft_agents.hosting.core import (
    AgentApplication,
    Authorization,
    MemoryStorage,
    MessageFactory,
    TurnContext,
    TurnState,
)

from .services.card import create_country_selection_card
from .services.message_route import handle_two_level_conversation_turn
from .services.retrieval_client import get_retrieval_options

logger = logging.getLogger(__name__)
load_dotenv()
configuration = load_configuration_from_env(environ)
storage = MemoryStorage()
connection_manager = MsalConnectionManager(**configuration)
adapter = CloudAdapter(connection_manager=connection_manager)
authorization = Authorization(storage, connection_manager, **configuration)
get_retrieval_options()
AGENT_APP = AgentApplication[TurnState](
    storage=storage, adapter=adapter, authorization=authorization, **configuration
)


@AGENT_APP.conversation_update("membersAdded")
async def welcome(context: TurnContext, _state: TurnState) -> None:
    for member in context.activity.members_added:
        if member.id != context.activity.recipient.id:
            await context.send_activity(
                MessageFactory.text(
                    "Hello! Welcome to the **Two-Level Retrieval with Microsoft 365 Copilot Retrieval API** demo. "
                    "Before searching SharePoint (`MyFiles` and `OrgFiles`), please select a **Country** (`Europe`, `US`, or `China`)."
                )
            )
            await context.send_activity(
                MessageFactory.attachment(
                    Attachment(
                        content_type="application/vnd.microsoft.card.adaptive",
                        content=create_country_selection_card(),
                    )
                )
            )


@AGENT_APP.message(re.compile(r".*", re.DOTALL), auth_handlers=["GRAPH"])
async def message(context: TurnContext, _state: TurnState) -> None:
    await context.send_activity(Activity(type=ActivityTypes.typing))

    async def get_access_token() -> str | None:
        token_response = await AGENT_APP.auth.get_token(context, "GRAPH")
        return token_response.token if token_response else None

    conversation_id = (
        context.activity.conversation.id
        if context.activity and context.activity.conversation
        else "default"
    )
    activity_value = (
        context.activity.value
        if isinstance(getattr(context.activity, "value", None), dict)
        else None
    )
    await handle_two_level_conversation_turn(
        conversation_id=conversation_id,
        text=context.activity.text,
        value=activity_value,
        get_access_token=get_access_token,
        send_text=lambda text: context.send_activity(MessageFactory.text(text)),
        send_card=lambda card: context.send_activity(
            MessageFactory.attachment(
                Attachment(
                    content_type="application/vnd.microsoft.card.adaptive",
                    content=card,
                )
            )
        ),
    )
