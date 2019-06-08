// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityTypes } = require('botbuilder');
const { QnAMaker } = require('botbuilder-ai');

/**
 * A simple bot that responds to utterances with answers from QnA Maker.
 * If an answer is not found for an utterance, the bot responds with help.
 */
const { ActivityHandler } = require('botbuilder');

class QnABot extends ActivityHandler {
    /**
     * The QnAMakerBot constructor requires one argument (`endpoint`) which is used to create an instance of `QnAMaker`.
     * @param {QnAMakerEndpoint} endpoint The basic configuration needed to call QnA Maker. In this sample the configuration is retrieved from the .bot file.
     * @param {QnAMakerOptions} config An optional parameter that contains additional settings for configuring a `QnAMaker` when calling the service.
     */
    constructor(endpoint, qnaOptions) {
        super();
        this.qnaMaker = new QnAMaker(endpoint, qnaOptions);

        /**
         * Every conversation turn for our QnABot will call this method.
         * There are no dialogs used, since it's "single turn" processing, meaning a single request and
         * response, with no stateful conversation.
         * @param {TurnContext} turnContext A TurnContext instance, containing all the data needed for processing the conversation turn.
         */
        this.onMessage(async turnContext => {
            // By checking the incoming Activity type, the bot only calls QnA Maker in appropriate cases.
            if (turnContext.activity.type === ActivityTypes.Message) {
                // Perform a call to the QnA Maker service to retrieve matching Question and Answer pairs.
                const qnaResults = await this.qnaMaker.getAnswers(turnContext);

                // If an answer was received from QnA Maker, send the answer back to the user.
                if (qnaResults[0]) {
                    await turnContext.sendActivity(qnaResults[0].answer);

                // If no answers were returned from QnA Maker, reply with help.
                } else {
                    await turnContext.sendActivity('No QnA Maker answers were found.');
                }
            }
        });
    }
}

module.exports.QnABot = QnABot;
