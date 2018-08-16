var restify          = require("restify");
var botBuilder       = require("botbuilder");
var teams            = require("botbuilder-teams");
var request          = require("request");
var jwtDecode       = require("jwt-decode");


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
var tenant;

bot.dialog('/', function(session){
    console.log(teams.TeamsMessage.getTenantId(session.message));
    tenant = teams.TeamsMessage.getTenantId(session.message);
    request.get("https://login.microsoftonline.com/{tenant}/oauth2/authorize?client_id=" + process.env.APP_CLIENT_ID +
                "&response_type=code&redirect_uri=" + encodeURI("https://test-teams-bot.herokuapp.com/verified") +
                "&response_mode=query&resource=" + encodeURI("https://factset.onmicrosoft.com/9fac8285-6f7b-4cab-b385-6ed8aec01fde" )+ "&state=verified", function(error, response, body){
                    console.log(body);
                });
    //session.send("Your bot is running. You said: %s", session.message.text);
});

server.get("/verified", function(req, response){
    if(req.params.error) console.log("Access Denied" + req.params.error);
    if(req.params.state !== "verified") {console.log("CSRF error");}
    else {
        var code = req.params.code;
        var options = {
            host: "https://login.microsoftonline.com/",
            path: tenant + "/oauth2/token",
            headers: {
                "Content-type":  "application/x-www-form-urlencoded",
                "grant_type"  :  "authorization_code",
                "client_id"   :  process.env.APP_CLIENT_ID,
                "code"        :  code,
                "redirect_uri":  encodeURI("https://test-teams-bot.herokuapp.com/verified"),
                "resource"    :  encodeURI("https://factset.onmicrosoft.com/9fac8285-6f7b-4cab-b385-6ed8aec01fde"),
            },
        }
        request.post(options, function(error, response, body){
            console.log(body);
            var responseToken = JSON.parse(body);
            var decoded = jwt_decode(responseToken["id_token"]);
            console.log(decoded["unique_name"]);
            bot.send("Your bot is running. Your name is %s", decoded["unique_name"]);
        })
    }
});

server.get("/", function(req, response){ response.send(200,"Your app is up and running.")})

server.get("/terms", function(req, response){ response.send(200, "Sample terms of Usage page")})

server.get("/privacy", function(req, response){ response.send(200, "Privacy policy")});
