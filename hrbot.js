// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');

const { ContactHR } = require('./componentDialogs/contactHR');

const { ContactITServices } = require('./componentDialogs/contactITServices');
const {CardFactory} = require('botbuilder');

const {LuisRecognizer, QnAMaker}  = require('botbuilder-ai');
const optionsCard = require('./resources/adaptiveCards/optionsCard4');

const CHOICE_PROMPT    = 'CHOICE_PROMPT';
const TEXT_PROMPT    = 'TEXT_PROMPT';

const CARDS = [optionsCard]


var configResultHeaderLiteral;
var numberOfresultsToShow;
var resultToBeShown = '';
var processingDept = "n";
const asteriskLine = "*************************************";
const peopleDept = "people";
const itServicesDept = "it services";

const chooseDomainIntent = "chooseDomain";
const askQuestionIntent = "askQuestion";
const cancelIntent = "cancelIntent";
const greetingIntent = "greetingIntent";
const noneIntent = "None";
const contactHRIntent = "contactHR";
const contactITServicesIntent = "contactITServices";
const doneIntent = "doneIntent";
const helpIntent = "helpIntent";
const msteams = "msteams";
const webchat = "webchat";
const emulator = "emulator";

const configMaxResults = 3;
const domainSelector = ["People", "IT Services", 'Help', 'Cancel'];
const domainSelectorJSON = {"Items": [
    {
        "label": "People (Human Resources)",
        "value": "People"
    },
    {
        "label": "IT Services (Policies and Procedures)",
        "value": "IT Services"
    },
    {
        "label": "Help - About TaihoBuddy",
        "value": "Help"
    },
    {
        "label": "Cancel / Done",
        "value": "Cancel"
    }

    ] 
}
const selectorITServices = ['Done',  'Ask another question', 'Contact IT Services'];

const selectorITServicesJSON = {"Items": [
    {
        "label": "I am done",
        "value": "Done"
    },
    {
        "label": "Ask another question",
        "value": "Ask another question"
    },
    {
        "label": "Contact IT Services",
        "value": "Contact IT Services"
    }

    ] 
}
const selectorPeople = ['Done', 'Ask another question', 'Contact People'];
const selectorPeopleJSON = {"Items": [
    {
        "label": "I am done",
        "value": "Done"
    },
    {
        "label": "Ask another question",
        "value": "Ask another question"
    },
    {
        "label": "Contact People",
        "value": "Contact People"
    }

    ] 
}

//const welcomeText = " - Welcome! I am ready to answer your query to the best of my ability. Please choose the department and ask your question...";
const greetingText = " - Hello! Please choose the department and ask your question...";
const chooseDepartmentText = "Sure. Please choose the department...";
const noResultText = "### Sorry, your search has yielded no result. Please try another search or contact ";
const byeText = "Bye now... just say Hello to wake me up again!";
const oneResultText = "# There is only one result: ";
const welcomeText = "Welcome to Taiho Buddy!! Please choose the department...";
const searchConfirmText1 = "Please type your question, we will search the " ;
const searchConfirmText2 = " Knowledge Base and get you the best results!" ;
const searchYieldText1 = "# Your search has yielded ";
const searchYieldText2 = " results: ";
const confidenceScoreText = "\n \n" + "**Confidence score:** "
//const helpText = "Hello, my name is TaihoBuddy! I can help you search for answers to your question. Please select the department and then just ask your question!"
const helpText = "## Hello! My name is TaihoBuddy! " +
" My job is to answer your queries by foraging the curated Knowledge Bases and get you the most relevant answers. " + 
" However, if you think that you need to contact a personnel at some point, then you can do so by choosing the department during the conversation, like - Contact People..." + 
" When you do so, the department will have the full context of your search and will contact you via eMail / Phone to resolve your question. ";

class hrbot extends ActivityHandler {
    constructor(conversationState,userState) {
        super();
        console.log (userState)

        this.conversationState = conversationState;
        console.log ("**************user state*************************")
        this.userState = userState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.contactHRDialog = new ContactHR(this.conversationState,this.userState);
        this.contactITServicesDialog = new ContactITServices(this.conversationState,this.userState);
        
        
        this.previousIntent = this.conversationState.createProperty("previousIntent");
        this.conversationData = this.conversationState.createProperty('conservationData');        

        const dispatchRecognizer = new LuisRecognizer({
            applicationId: process.env.LuisAppId,
            endpointKey: process.env.LuisAPIKey,
            endpoint: `https://${ process.env.LuisAPIHostName }.api.cognitive.microsoft.com`
        }, {
            includeAllIntents: true
        }, true);

       
        const qnaMaker = new QnAMaker({
            knowledgeBaseId: process.env.QnAKnowledgebaseId,
            endpointKey: process.env.QnAEndpointKey,
            host: process.env.QnAEndpointHostName
        });

        const qnaMaker2 = new QnAMaker({
            knowledgeBaseId: process.env.QnAKnowledgebaseId,
            endpointKey: process.env.QnAEndpointKey,
            host: process.env.QnAEndpointHostName
        });
              
        this.qnaMaker = qnaMaker;
        this.qnaMaker2 = qnaMaker2;


        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            console.log ("In onMessage " )
            const luisResult = await dispatchRecognizer.recognize(context)
            const intent = LuisRecognizer.topIntent(luisResult);    
            const entities = luisResult.entities;
            await this.dispatchToIntentAsync(context,intent,entities);        
            await next();
        });

        this.onDialog(async (context, next) => {
            console.log ("In onDialog ")
           // console.log (JSON.stringify(context))
            // Save any state changes. The load happened during the execution of the Dialog.
            await this.conversationState.saveChanges(context, false);
            await this.userState.saveChanges(context, false);
            await next();
        });   

        this.onMembersAdded(async (context, next) => {
            console.log ("In onMembersAdded " )
            await this.sendWelcomeMessage(context)
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

  

    async sendWelcomeMessage(turnContext) {
        const { activity } = turnContext;

        // Iterate over all new members added to the conversation.
        for (const idx in activity.membersAdded) {
            if (activity.membersAdded[idx].id !== activity.recipient.id) {
               // const welcomeMessage = `Welcome to People Buddy ${ activity.membersAdded[idx].name }. Please choose the department`;
               var channelId = turnContext._activity.channelId
               console.log ("In welcome " + channelId)
                if (channelId === msteams){
                    var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
                } else {                
                    var cardGen = generateAdaptiveCard(domainSelectorJSON)
                }
               //console.log ("CARDS " + JSON.stringify(CARDS[0]))
               

               cardGen = JSON.parse(cardGen)
            
               var CARDS2 = [cardGen];
               await turnContext.sendActivity({
                    text: welcomeText,
                    attachments: [CardFactory.adaptiveCard(CARDS2[0])]
               });
              // const flow = await this.conversationFlow.get(turnContext, { lastQuestionAsked: question.none });
               
              // return await turnContext.prompt(TEXT_PROMPT, '');

            }
        }
    }

    async sendSuggestedActions(turnContext, selector) {        
        var reply = MessageFactory.suggestedActions(selector, "Please choose:");
        await turnContext.sendActivity(reply);
    }


    async dispatchToIntentAsync(context,intent,entities){
        console.log ("In dispatchToIntentAsync: " + intent)
        var channelId = context._activity.channelId
        const conversationData = await this.conversationData.get(context,{}); 
        var currentIntent = '';
        var QnAMakerOptions = {
            top:3
        }
        if(intent == chooseDomainIntent ){
            console.log ("getting department...")
            var dept = entities.department[0]
            console.log ("Department chosen: "+ dept)
            await this.conversationData.set(context,{deptSaved: dept});
            await context.sendActivity(searchConfirmText1 + dept.toUpperCase() + searchConfirmText2);
        }
        if(intent == askQuestionIntent ){
            console.log ("in askQuestion intent");
            await context.sendActivity(chooseDepartmentText);          
            //var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
            if (channelId === msteams){
                var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
            }else {
                 var cardGen = generateAdaptiveCard(domainSelectorJSON)
            }
            cardGen = JSON.parse(cardGen)
            var CARDS2 = [cardGen];
            await context.sendActivity({
                 text: '',
                 attachments: [CardFactory.adaptiveCard(CARDS2[0])]
            }); 
        }
        if(intent == greetingIntent ){
            console.log ("in greetingIntent intent");
            console.log ("Channelid " + JSON.stringify(context._activity.channelId))
            
           // await context.sendActivity(greetingText);            
           // var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
            if (channelId === msteams){
                var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
            } else {
                var cardGen = generateAdaptiveCard(domainSelectorJSON)
            }
            cardGen = JSON.parse(cardGen)
            var CARDS2 = [cardGen];
            await context.sendActivity({
                 text: greetingText,
                 attachments: [CardFactory.adaptiveCard(CARDS2[0])]
            }); 
        }
        if((intent == helpIntent)){
            console.log ("in help intent");
            const conversationData = await this.conversationData.get(context,{});  
            console.log ("processingDept " + conversationData.processingDept)
            console.log ("contactPeopleDone " + conversationData.contactPeopleDone)
            
            await context.sendActivity(helpText);  
           // var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
            if (channelId === msteams){
                var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
            } else {
                var cardGen = generateAdaptiveCard(domainSelectorJSON)
            }
            cardGen = JSON.parse(cardGen)
            var CARDS2 = [cardGen];
            await context.sendActivity({
                 text: '',
                 attachments: [CardFactory.adaptiveCard(CARDS2[0])]
            });          
           // await this.sendSuggestedActions(context, domainSelector);
        }
        if(intent == noneIntent ){
            console.log ("In none intent, calling QNA Maker")
            const conversationData = await this.conversationData.get(context,{});  
            console.log (conversationData.deptSaved)
            var selectorDialog;
            var result;
            if (conversationData.deptSaved === peopleDept){
                console.log("searching in People Knowledge Base")
                selectorDialog = selectorPeople
                result = await this.qnaMaker.getAnswers(context, QnAMakerOptions)
            }

            if (conversationData.deptSaved === itServicesDept){
                console.log("searching in IT Services Knowledge Base")
                selectorDialog = selectorITServices
                result = await this.qnaMaker.getAnswers(context, QnAMakerOptions)
            }

            console.log ("***************************************")
            console.log ("Number of rows returned: " + JSON.stringify(result.length))
            console.log ("***************************************")

            //Handle max results to show
            if (result.length > 0){
                var numberOffResultsReturned = result.length
                if (configMaxResults > numberOffResultsReturned){
                    numberOfresultsToShow  = numberOffResultsReturned
                } else{
                    numberOfresultsToShow = configMaxResults
                }
                if (numberOfresultsToShow === 1){
                    configResultHeaderLiteral = oneResultText;
    
                } else{
                    configResultHeaderLiteral = searchYieldText1 + numberOfresultsToShow + searchYieldText2
                }
                

                console.log("configMaxResults      " + configMaxResults)
                console.log("numberOffResultsReturned " + numberOffResultsReturned)
                console.log("numberOfresultsToShow " + numberOfresultsToShow)
            }

            if (result.length > 0){
                
                resultToBeShown = ''
                for (var i=0; i<numberOfresultsToShow; i++){
                    var score = (Math.round(result[i].score * 100) / 100).toFixed(2);
                    var resultnumber = "## Result [" + (i+1) + "]"
                    resultToBeShown =  resultToBeShown + "\n \n" + resultnumber + "\n \n" + result[i].answer + confidenceScoreText + score + "\n \n" +  "**Source:** "  + result[i].source   + "\n \n" + asteriskLine
                }
                await context.sendActivity(configResultHeaderLiteral + "\n \n" + asteriskLine + "\n \n" + resultToBeShown);

            }  else{
                await context.sendActivity(noResultText + conversationData.deptSaved.toUpperCase() + " department");
            }
            if (conversationData.deptSaved === itServicesDept){
                console.log("selecting department new IT SErvices")
                //selectorDialog = selectorITServices  
              //  var cardGen = generateAdaptiveCardTeams(selectorITServicesJSON)
                if (channelId === msteams){
                    var cardGen = generateAdaptiveCardTeams(selectorITServicesJSON)
                } else {
                    var cardGen = generateAdaptiveCard(selectorITServicesJSON)
                }
                cardGen = JSON.parse(cardGen)
                var CARDS2 = [cardGen];
                await context.sendActivity({
                     text: '',
                     attachments: [CardFactory.adaptiveCard(CARDS2[0])]
                });              
            }
            if (conversationData.deptSaved === peopleDept){
                console.log("selecting department new People ")
                //selectorDialog = selectorPeople
               // var cardGen = generateAdaptiveCardTeams(selectorPeopleJSON)
                if (channelId === msteams){
                    var cardGen = generateAdaptiveCardTeams(selectorPeopleJSON)
                } else {
                    var cardGen = generateAdaptiveCard(selectorPeopleJSON)
                }
                cardGen = JSON.parse(cardGen)
                var CARDS2 = [cardGen];
                await context.sendActivity({
                     text: '',
                     attachments: [CardFactory.adaptiveCard(CARDS2[0])]
                }); 
                
                
            }
                        
           // await this.sendSuggestedActions(context, selectorDialog);
        }
        else
        {
            currentIntent = intent;
            console.log ("currentIntent here, yet to decide department: " + currentIntent)
            const conversationData = await this.conversationData.get(context,{}); 
            console.log ("conversationData.contactPeopleDone " + JSON.stringify(conversationData)) 


            if (conversationData.contactPeopleDone === false){
                console.log ("Forcing intent to stick to conversation")
                currentIntent = contactHRIntent
            }
            if (currentIntent === contactHRIntent){
                console.log ("In contactHR intent");
                console.log ("setting contactPeopleDone to false")
                await this.conversationData.set(context,{endDialog: false, contactPeopleDone: false, processingDept: true}); 
                             
                await this.contactHRDialog.run(context,this.dialogState,entities);
                conversationData.endDialog = await this.contactHRDialog.isDialogComplete();
                console.log (conversationData.endDialog);
                if(conversationData.endDialog)
                {
                    await this.conversationData.set(context,{endDialog: true, processingDept: false, contactPeopleDone: true}); 
                   // await this.previousIntent.set(context,{intentName: null}); 
                } 
            } else

            if (currentIntent === contactITServicesIntent){
                console.log ("In intent contactITServices")
                console.log ("setting contactITServicesDone to false")
                await this.conversationData.set(context,{endDialog: false,contactITServicesDone: false , processingDept: true});
                await this.contactITServicesDialog.run(context,this.dialogState,entities);
                conversationData.endDialog = await this.contactITServicesDialog.isDialogComplete();
                if(conversationData.endDialog)
                {
                    await this.conversationData.set(context,{endDialog: true, contactITServicesDone: true, processingDept: false});
                   // await this.previousIntent.set(context,{intentName: null}); 
                } 

            }
            if  ((intent == doneIntent) || (intent == cancelIntent)){
                console.log ("In done  / cancel intent " + JSON.stringify(intent))
                await context.sendActivity(byeText);
                
            }

        }
    
    }
}


 function generateAdaptiveCard(jsonObject) {


    var len  = jsonObject.Items.length
    var jsonStr = '';

    for (var i=0;i<len;i++){
         jsonStr = jsonStr +'{ '
        + '"type" : "Action.Submit"'       
        + ", "
        + '"title" : '
        + '"'+ jsonObject.Items[i].label + '"'
        + ", "
        + '"data" : '
        + '"'+ jsonObject.Items[i].value + '"' 
        + ' }'
        + ", ";         
    }
    //Remove trailing comma
    jsonStr = jsonStr.replace(/,\s*$/, "");
    
    var cardFormatted = '{ '
    + '"$schema": "https://adaptivecards.io/schemas/adaptive-card.json"'
    + ", "
    + '"type": "AdaptiveCard"'
    + ", "
    + '"version": "1.0"'
    + ", "
    + '"actions":'
    + " [ " 
    +  jsonStr 
    +  "]"
    + "}"

    
    return cardFormatted
}
function generateAdaptiveCardTeams(jsonObject) {


    var len  = jsonObject.Items.length
    var jsonStr = '';

    for (var i=0;i<len;i++){
         jsonStr = jsonStr +'{ '
        + '"type" : "Action.Submit"'       
        + ', '
        + '"title" : '
        + '"'+ jsonObject.Items[i].label + '"'
        + ', '
        + '"data" : '
        + '{'
        + '"msteams" :' 
        + '{'
        + '"type": "messageBack"'
        + ', '
        + '"displayText": ' 
        + '"'+ jsonObject.Items[i].label + '"' +  ', '
        + '"text": ' 
        + '"'+ jsonObject.Items[i].value + '"' +  ', '
        + '"value": '
        + '"'+ jsonObject.Items[i].value + '"' 

       
        + ' }'
        + ' }'
        + ' }'
        + ", ";         
    }
    //Remove trailing comma
    jsonStr = jsonStr.replace(/,\s*$/, "");
    
    var cardFormatted = '{ '
    + '"$schema": "https://adaptivecards.io/schemas/adaptive-card.json"'
    + ", "
    + '"type": "AdaptiveCard"'
    + ", "
    + '"version": "1.0"'
    + ", "
    + '"actions":'
    + " [ " 
    +  jsonStr 
    +  "]"
    + "}"

    
    return cardFormatted
}




module.exports.hrbot = hrbot;
