var restify          = require("restify");
var botBuilder       = require("botbuilder");
var teams            = require("botbuilder-teams");


var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3980, function () {
    console.log('%s listening to %s', server.name, server.url); 
 });

var connector = new botBuilder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,// || botConfig.microsoftAppId,
    appPassword: process.env.MICROSOFT_APP_PASSWORD// || botConfig.microsoftAppPassword
});

server.post('/api/messages',connector.listen());

var inMemoryStorage = new botBuilder.MemoryBotStorage();

var bot = new botBuilder.UniversalBot(connector).set('storage', inMemoryStorage);

bot.dialog('/', function(session){
    console.log(teams.TeamsMessage.getTenantId(session.message));
    session.send("Your bot is running. You said: %s", session.message.text);
})


server.get("/", function(req, response){ response.send(200,"Your app is up and running.")})

server.get("/terms", function(req, response){ response.send(200, "Sample terms of Usage page")})

server.get("/privacy", function(req, response){ response.send(200, "Privacy policy")})


































































