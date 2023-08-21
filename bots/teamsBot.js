// <copyright file="teamsBot.js" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

const {TeamsActivityHandler, CardFactory} = require("botbuilder");

class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();
    }

    // Invoked when an app based link query activity is received from the connector.
    async handleTeamsAppBasedLinkQuery(context, query) {
        const userCard = this.getLinkUnfurlingCard(context.activity.from.id);

        return {
            composeExtension: {
                type: 'result',
                attachmentLayout: 'list',
                attachments: [{
                    contentType: "application/vnd.microsoft.card.adaptive",
                    content: userCard,
                    preview: {
                        contentType: "application/vnd.microsoft.card.adaptive",
                        content: userCard
                    }
                }],
                suggestedActions: {
                    actions: [
                        {
                            type: "setCachePolicy",
                            value: {type: "no-cache"}
                        }
                    ]
                }
            }
        }
    }

    async onAdaptiveCardInvoke(context, invokeValue) {
        const card = {
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.5",
            "height": "auto",
            "body": [
                {
                    "type": "TextBlock",
                    "size": "extraLarge",
                    "text": "Adaptive Card",
                    "height": "auto"
                },
                {
                    "type": "Image",
                    "url": "https://raw.githubusercontent.com/microsoft/botframework-sdk/master/icon.png",
                    "height": "auto"
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Additional action",
                    "data": {
                        "actionName": "actionValue"
                    }
                }
            ]
        };

        return {
            statusCode: 200,
            type: "application/vnd.microsoft.card.adaptive",
            value: card
        };
    }

    // Adaptive card for link unfurling.
    getLinkUnfurlingCard = (senderId) => {
        return {
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.5",
            "height": "auto",
            "body": [
                {
                    "type": "TextBlock",
                    "size": "extraLarge",
                    "text": "Adaptive Card",
                    "height": "auto"
                },
                {
                    "type": "Image",
                    "url": "https://raw.githubusercontent.com/microsoft/botframework-sdk/master/icon.png",
                    "height": "auto"
                }
            ],
            "refresh": {
                "action": {
                    "type": "Action.Execute",
                    "title": "Submit",
                    "data": {
                        "someData": "someValue"
                    }
                },
                "userIds": [senderId]
            },
            "actions": []
        }
    }
}

module.exports.TeamsBot = TeamsBot;