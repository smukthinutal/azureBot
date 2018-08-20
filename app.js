var restify          = require("restify");
var botBuilder       = require("botbuilder");
var teams            = require("botbuilder-teams");
var request          = require("request");
var jwtDecode        = require("jwt-decode");
var rn               = require('random-number');


var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3980, function () {
    console.log('%s listening to %s', server.name, server.url); 
 });

var connector = new botBuilder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,// || botConfig.microsoftAppId,
    appPassword: process.env.MICROSOFT_APP_PASSWORD// || botConfig.microsoftAppPassword
});

var tenantId;

server.post('/api/messages',connector.listen());

var inMemoryStorage = new botBuilder.MemoryBotStorage();

var userKey = {};

var bot = new botBuilder.UniversalBot(connector).set('storage', inMemoryStorage);

var tempSession;

var csrfRandomOptions = { min: 0, max: 100000, integer: true};
var generateRandom = rn.generator(csrfRandomOptions);
var csrfRandomNumber;

bot.dialog('/', function(session){
    tenantId = teams.TeamsMessage.getTenantId(session.message);
    if(session.userData.accessKey) {
        session.send("Hi %s, you are already logged in", session.message.user.name);
        tempSession = session;
    }
    else if( userKey[session.message.user.name]) {
        session.userData.accessKey = userKey[session.message.user.name];
        session.send("You have been successfully authenticated");
    }
    else {
        session.send("You have to [login](https://test-teams-bot.herokuapp.com/login) before using this bot");
    }
});

server.get("/login",function(req, res, next){
    csrfRandomNumber = generateRandom();
    var loginURL = "https://login.microsoftonline.com/" + tenantId + "/oauth2/authorize?client_id=" + process.env.APP_CLIENT_ID +
                   "&response_type=code&redirect_uri=" + encodeURIComponent( "https://" + req.headers.host + "/verified") +
                   "&response_mode=query&resource=" + encodeURIComponent("https://graph.microsoft.com")+ "&state=" + csrfRandomNumber;
   console.log(loginURL);
   res.redirect(loginURL, next);
});

server.use(restify.plugins.queryParser());
server.get("/verified", function(req, res){
   if(req.query.state !== csrfRandomNumber) res.send(401, "CSRF error");
   else {
       var authURLOptions = {
           host: "https://login.microsoftonline.com/",
           path: tenantId + "/oauth2/token",
           headers: {
               "Content-type":  "application/x-www-form-urlencoded",
           },
           body: "grant_type=authorization_code&client_id=" + process.env.APP_CLIENT_ID + "&code="+ req.query.code + "&redirect_uri=" + 
               encodeURIComponent("https://" + req.headers.host + "/verified") + "&resource=" + encodeURIComponent("https://graph.microsoft.com") +
               "&client_secret=" + encodeURIComponent(process.env.APP_KEY)
       }
       request.post("https://login.microsoftonline.com/" + tenantId + "/oauth2/token", authURLOptions, function(error, response, body){
           console.log(error);
           console.log(body);
           var json = JSON.parse(body);
           var decoded = jwtDecoder(json["id_token"]);
           console.log(decoded);
           var getOptions = {
               headers: {
                   "Content-Type"  : "application/json",
                   "Authorization" : json["access_token"]
               }
           }
           var details = request.get("https://graph.microsoft.com/v1.0/me", getOptions, function(getError, getResponse, getBody){
               if(getResponse.statusCode !== 200) {
                   console.log("Refreshing token for user: " + decoded["unique_name"] );
                   var refreshOptions = {
                       host: "https://login.microsoftonline.com",
                       headers: {
                           "Content-type":  "application/x-www-form-urlencoded",
                       },
                       body:   "grant_type=refresh_token&client_id=" + process.env.APP_CLIENT_ID +
                               "&refresh_token=" + json["refresh_token"] +
                               "&resource=" + encodeURIComponent("https://graph.microsoft.com") +
                               "&client_secret=" + encodeURIComponent(process.env.APP_KEY)
                   }
                   request.post("https://login.microsoftonline.com/" + tenantId + "/oauth2/token", refreshOptions, function(refreshError, refreshResponse, refreshBody){
                       console.log(refreshError);
                      // console.log(refreshBody);
                       var refreshJson = JSON.parse(refreshBody);
                       getOptions.headers.Authorization = refreshJson.access_token;
                       if(refreshJson.error) console.log(refreshJson.error);
                       else {
                           request.get("https://graph.microsoft.com/v1.0/me", getOptions, function(refreshGetError, refreshGetResponse, refreshGetBody){
                              // console.log(refreshGetBody);
                               console.log(refreshGetError);
                           });
                       }
                   });
               }
               else {
                   console.log(getError);
                   //console.log(getBody);
               }
           });
           userKey[decoded.name] = "authenticated";
           tempSession.beginDialog("/");
           res.send(200, "Successfully authenticated");
       });
   }
});

server.get("/", function(req, response){ response.send(200,"Your app is up and running.")})

server.get("/terms", function(req, response){ response.send(200, "Sample terms of Usage page")})

server.get("/privacy", function(req, response){ response.send(200, "Privacy policy")});
