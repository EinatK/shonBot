

var restify = require('restify');
var builder = require('botbuilder');
var XLSX = require('xlsjs');

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 

var connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

server.post('/api/messages', connector.listen());


var MainActions = {
    Action1: 'action a',
    Action2: 'action b',
    Action3: 'action c',
    Action4: 'action d',
    Action5: 'action e',
};

/*
Seems like doing it with json is much better then creating an array but we'll see itll take more time
function generateClientArray (){
    var clients = XLSX.utils.sheet_to_json('/Users/einatkidron/Downloads/botData.xlsx');
}*/

function getClientName(cellAdd) {
    var workbook = XLSX.readFile('/Users/einatkidron/Downloads/botData.xlsx');    
    
    var clientData = workbook.SheetNames[0];
    var worksheet = workbook.Sheets[clientData]; 
    var address_of_cell = cellAdd;
    var desired_cell = worksheet[address_of_cell];         
    var desired_value = (desired_cell ? desired_cell.v : undefined);

      
    return desired_value; 
  }


var bot = new builder.UniversalBot(connector, function (session) {

    //The card for the client names - to be read from an excel that should be parsed into a json
    var clientsCard = new builder.HeroCard(session)    
    .title('Hello')
    .subtitle('Please choose a client')
    .buttons([
        builder.CardAction.imBack(session, session.gettext(getClientName('A1')), getClientName('A1')),
        builder.CardAction.imBack(session, session.gettext(getClientName('A2')), getClientName('A2')),
        
    ]);

    session.send(new builder.Message(session)
    .addAttachment(clientsCard));



    /*MISSING: implementation of waiting for a response before showing the second card. There's a better way 
    to do this, check out this link, it waits for a response too
    http://blog.geektrainer.com/2017/06/08/Working-with-custom-buttons-to-drive-conversations/
    */



    //the card of the actions - taken from a local dict above
    var actionsCard = new builder.HeroCard(session)
    .title('What would you like to do?')
    .subtitle('Pick an action')
    .buttons([
        builder.CardAction.imBack(session, session.gettext(MainActions.Action1), MainActions.Action1),
        builder.CardAction.imBack(session, session.gettext(MainActions.Action2), MainActions.Action2),
        builder.CardAction.imBack(session, session.gettext(MainActions.Action3), MainActions.Action3),
        builder.CardAction.imBack(session, session.gettext(MainActions.Action4), MainActions.Action4),
        builder.CardAction.imBack(session, session.gettext(MainActions.Action5), MainActions.Action5),    
    ])

    session.send(new builder.Message(session)
    .addAttachment(actionsCard));

    session.endDialogWithResult('Thank you');
});


})