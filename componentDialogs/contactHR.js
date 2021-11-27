const {WaterfallDialog, ComponentDialog } = require('botbuilder-dialogs');
const { ActivityHandler, MessageFactory } = require('botbuilder');

const {ConfirmPrompt, ChoicePrompt, DateTimePrompt, NumberPrompt, TextPrompt  } = require('botbuilder-dialogs');

const {DialogSet, DialogTurnStatus } = require('botbuilder-dialogs');
const {CardFactory} = require('botbuilder');


const CHOICE_PROMPT    = 'CHOICE_PROMPT';
const CONFIRM_PROMPT   = 'CONFIRM_PROMPT';
const TEXT_PROMPT      = 'TEXT_PROMPT';
const NUMBER_PROMPT    = 'NUMBER_PROMPT';
const DATETIME_PROMPT  = 'DATETIME_PROMPT';
const ACTIVITY_PROMPT  = 'ACTIVITY_PROMPT';
const WATERFALL_DIALOG = 'WATERFALL_DIALOG';

const domainSelector = ["People", "IT Services", 'Cancel'];
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
const problemAreaPeople = ["Benefits", "Covid", "Training", "Vacation", "Cancel"];
const problemAreaPeopleJSON = {"Items": [
    {
        "label": "Wanted to know about Benefits",
        "value": "Benefits"
    },
    {
        "label": "Information on Covid",
        "value": "Covid"
    },
    {
        "label": "Training related...",
        "value": "Training"
    },
    {
        "label": "Vacation policy related",
        "value": "Vacation"
    },
    {
        "label": "Cancel",
        "value": "Cancel"
    }

    ] 
}
const problemBriefOptions= ["Results not useful", "Need more info", "No Results", "Timed out", "Cancel"];
const problemBriefOptionsJSON = {"Items": [
    {
        "label": "Results were not useful",
        "value": "Results not useful"
    },
    {
        "label": "Need more information",
        "value": "Need more info"
    },
    {
        "label": "Search returned no result",
        "value": "No Results"
    },
    {
        "label": "Search Timed out!",
        "value": "Timed out"
    },
    {
        "label": "Cancel",
        "value": "Cancel"
    }

    ] 
}
const problemAreaText = "Please select the topic about which you need assistance...";
const problemBriefText = "And what went wrong?";
const emailSentText = " \n \n eMail sent to People Team. They will get back to ASAP. You can continue with your search..."
var endDialog ='';

class ContactHR extends ComponentDialog {
    
    constructor(conversationState,userState) {
        super('contactHR');

        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new NumberPrompt(NUMBER_PROMPT,this.noOfParticipantsValidator));
        this.addDialog(new DateTimePrompt(DATETIME_PROMPT));
        this.conversationState = conversationState;
        this.conversationData = this.conversationState.createProperty('conservationData');

        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
            this.getProblemArea.bind(this),  // Get getProblemArea           
            this.getProblemBrief.bind(this),    // Get getProblemBrief
            this.sendEmail.bind(this),    // send Email            
        ]));

        this.initialDialogId = WATERFALL_DIALOG;

   }

    async run(turnContext, accessor, entities) {
        console.log ("in run...")
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        
        const results = await dialogContext.continueDialog();
       
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id, entities);
        }
    }

    async getProblemArea(step) {
        console.log ("In getProblemArea");
        step.values.contactPeopleDone = false  
        endDialog = false;
        // Running a prompt here means the next WaterfallStep will be run when the users response is received.
        var cardGen = generateAdaptiveCard(problemAreaPeopleJSON)
        cardGen = JSON.parse(cardGen)
        var CARDS2 = [cardGen];
        var greetingText = "Please choose the problem area..."

        const cardPrompt = MessageFactory.text('');
        await step.context.sendActivity({
            text: greetingText,
            attachments: [CardFactory.adaptiveCard(CARDS2[0])]
       });
       return await step.prompt(TEXT_PROMPT, 'Tip: Problem area is the sub-domain within the People department about which you had raised a query');
 
           
    }

    async getProblemBrief(step){
        console.log ("In getProblemBrief")  
        step.values.contactPeopleDone = false      
        step.values.probArea = step.result
        //console.log ("step.result " + step.result)
        console.log ("step.values.probArea " + step.values.probArea)
        var cardGen = generateAdaptiveCard(problemBriefOptionsJSON)
        cardGen = JSON.parse(cardGen)
        var CARDS2 = [cardGen];
        var greetingText = "And the problem you faced..."
        await step.context.sendActivity({
             text: greetingText,
             attachments: [CardFactory.adaptiveCard(CARDS2[0])]
        }); 
        return await step.prompt(TEXT_PROMPT, 'Tip: Problem brief is the reason due to which you considered contacting the department');
      //  return await step.prompt(CHOICE_PROMPT, problemBriefText, problemBriefOptions);        
    }

    async sendEmail(step){
        console.log ("In sendEmail") 
        console.log (step.values.probArea)
        var probBrief = step.result;
        await step.context.sendActivity("### Problem Area: " + step.values.probArea + " ,  Problem brief: " + probBrief + emailSentText)
       // await this.sendSuggestedActions(step.context, domainSelector);
        var cardGen = generateAdaptiveCard(domainSelectorJSON)
        cardGen = JSON.parse(cardGen)
        var CARDS2 = [cardGen];
        var greetingText = "You can continue with search..."
        await step.context.sendActivity({
            // text: greetingText,
             attachments: [CardFactory.adaptiveCard(CARDS2[0])]
        }); 
        step.values.contactPeopleDone = true  
        endDialog = true;
        return await step.endDialog();   
    
    }
    async sendSuggestedActions(turnContext, selector) {
        var reply = MessageFactory.suggestedActions(selector);
        await turnContext.sendActivity(reply);
    }



    async isDialogComplete(){
        return endDialog;
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

module.exports.ContactHR = ContactHR;








