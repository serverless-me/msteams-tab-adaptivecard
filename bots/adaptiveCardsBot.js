// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TeamsActivityHandler, CardFactory, TurnContext, TeamsInfo } = require('botbuilder');
const { AdaptiveCard, TextBlock, HostConfig } = require("adaptivecards");
const { Template } = require("adaptivecards-templating");

// Import Template & Data
const templateJson = require('../resources/CardTemplate.json');
const dataJson = require('../resources/CardData.json');

const WELCOME_TEXT = 'This bot will introduce you to Adaptive Cards. Type anything to see an Adaptive Card.';

class AdaptiveCardsBot extends TeamsActivityHandler {
    constructor() {
        super();        
    }

    handleTeamsTabFetch(context, tabRequest) {
        const template = new Template(templateJson);
    
        const cardPayload = template.expand(dataJson);
        let card = new AdaptiveCard();
        card.parse(cardPayload);

        let jsonCard = card.toJSON();
        
        return {
            "tab": {
                "type": "continue",
                "value": {
                    "cards": [
                        {
                            "card": jsonCard,
                        } 
                    ]
                },
            },
            "responseType": "tab"
        };
    }
}

module.exports.AdaptiveCardsBot = AdaptiveCardsBot;
