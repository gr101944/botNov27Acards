const {WaterfallDialog, ComponentDialog } = require('botbuilder-dialogs');
const { ActivityHandler, MessageFactory } = require('botbuilder');
const {CardFactory} = require('botbuilder');
const msteams = "msteams";
const webchat = "webchat";
const emulator = "emulator";

const {ConfirmPrompt, ChoicePrompt, DateTimePrompt, NumberPrompt, TextPrompt  } = require('botbuilder-dialogs');

const {DialogSet, DialogTurnStatus } = require('botbuilder-dialogs');


const CHOICE_PROMPT    = 'CHOICE_PROMPT';
const CONFIRM_PROMPT   = 'CONFIRM_PROMPT';
const TEXT_PROMPT      = 'TEXT_PROMPT';
const NUMBER_PROMPT    = 'NUMBER_PROMPT';
const DATETIME_PROMPT  = 'DATETIME_PROMPT';
const WATERFALL_DIALOG = 'WATERFALL_DIALOG';
var endDialog ='';
const emailSentText = " \n \n eMail sent to IT Services Team. They will get back to ASAP. You can continue with your search..."
var domainSelector = ["People", "IT Services", 'Cancel']
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
var problemBriefOptions= ["Results not useful", "Need more info", "No Results", "Timed out", "Cancel"]
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
var problemAreaITServices = ["Equipment", "Policies", "Access Related", "Software", "Cancel"]
const problemAreaITServicesJSON = {"Items": [
    {
        "label": "About my equipment",
        "value": "Equipment"
    },
    {
        "label": "Regarding Policies",
        "value": "Policies"
    },
    {
        "label": "Access Related",
        "value": "Access Related"
    },
    {
        "label": "Software related",
        "value": "Software"
    },
    {
        "label": "Cancel",
        "value": "Cancel"
    }

    ] 
}


class ContactITServices extends ComponentDialog {
    
    constructor(conservsationState,userState) {
        super('contactITServices');

        this.addDialog(new TextPrompt(TEXT_PROMPT));
        this.addDialog(new ChoicePrompt(CHOICE_PROMPT));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
        this.addDialog(new NumberPrompt(NUMBER_PROMPT,this.noOfParticipantsValidator));
        this.addDialog(new DateTimePrompt(DATETIME_PROMPT));

        this.addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
            this.getProblemArea.bind(this),  // Get getProblemArea
            this.getProblemBrief.bind(this),    // Get getProblemBrief
            this.sendEmail.bind(this),    // send Email            
        ]));

        this.initialDialogId = WATERFALL_DIALOG;

   }

    async run(turnContext, accessor, entities) {
        var channelId = turnContext._activity.channelId
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id, entities);
        }
    }

    async getProblemArea(step, channelId) {
        
        console.log ("In getProblemArea");
        step.values.contactITServicesDone = false;
        console.log ("contactITServicesDone " + step.values.contactITServicesDone);
        endDialog = false;
        // Running a prompt here means the next WaterfallStep will be run when the users response is received.
       // var cardGen = generateAdaptiveCardTeams(problemAreaITServicesJSON)
        var channelId = step.context.activity.channelId;
        if (channelId === msteams){
            var cardGen = generateAdaptiveCardTeams(problemAreaITServicesJSON)
        } else {
            var cardGen = generateAdaptiveCard(problemAreaITServicesJSON)
        }
        cardGen = JSON.parse(cardGen)
        var CARDS2 = [cardGen];
        var greetingText = "Please choose the problem area..."

        const cardPrompt = MessageFactory.text('');
        await step.context.sendActivity({
            text: greetingText,
            attachments: [CardFactory.adaptiveCard(CARDS2[0])]
       });
       return await step.prompt(TEXT_PROMPT, 'Tip: Problem area is the sub-domain within the IT Services department about which you had raised a query');
           
    }

    async getProblemBrief(step, channelId){
        console.log ("In getProblemBrief")        
       // console.log(step.result)
        step.values.probArea = step.result
        step.values.contactITServicesDone = false
        console.log ("contactITServicesDone " + step.values.contactITServicesDone)
        
        //var cardGen = generateAdaptiveCardTeams(problemBriefOptionsJSON)
        var channelId = step.context.activity.channelId;
        if (channelId === msteams){
            var cardGen = generateAdaptiveCardTeams(problemBriefOptionsJSON)
        } else {
            var cardGen = generateAdaptiveCard(problemBriefOptionsJSON)
        }
        cardGen = JSON.parse(cardGen)
        var CARDS2 = [cardGen];
        var greetingText = "And the problem you faced..."
        await step.context.sendActivity({
             text: greetingText,
             attachments: [CardFactory.adaptiveCard(CARDS2[0])]
        }); 
        return await step.prompt(TEXT_PROMPT, 'Tip: Problem brief is the reason due to which you considered contacting the department');
        
        
    }

    async sendEmail(step, channelId){
        console.log ("In sendEmail");
        var probBrief = step.result;
        await step.context.sendActivity("### Problem Area: " + step.values.probArea + " ,  Problem brief: " + probBrief + emailSentText)
       // await this.sendSuggestedActions(step.context, domainSelector);
        //var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
        var channelId = step.context.activity.channelId;
        if (channelId === msteams){
            var cardGen = generateAdaptiveCardTeams(domainSelectorJSON)
        } else {      
            var cardGen = generateAdaptiveCard(domainSelectorJSON)
        }
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

module.exports.ContactITServices = ContactITServices;








