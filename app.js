var restify          = require("restify");
var botBuilder       = require("botbuilder");
var teams            = require("botbuilder-teams");
var request          = require("request");
var jwtDecoder       = require("jwt-decode");
var rn               = require('random-number');
var jsonwebtoken     = require("jsonwebtoken");


var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3980, function () {
    console.log('%s listening to %s', server.name, server.url); 
 });

var connector = new botBuilder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,// || botConfig.microsoftAppId,
    appPassword: process.env.MICROSOFT_APP_PASSWORD// || botConfig.microsoftAppPassword
});

var tenantId;


var inMemoryStorage = new botBuilder.MemoryBotStorage();

var bot = new botBuilder.UniversalBot(connector).set('storage', inMemoryStorage);

server.use(restify.plugins.bodyParser());
server.post('/api/messages',function(req, res, next){
    console.log(req.headers);
    var decoded = jsonwebtoken.decode(req.headers.authorization.replace("Bearer ",""));
    console.log( "Header:" + decoded.header);
    console.log("payload:" + decoded.payload);
    console.log(req.body);
    connector.listen();
});

var userKey = {};

var tempAddress;

var csrfRandomOptions = { min: 0, max: 100000, integer: true};
var generateRandom = rn.generator(csrfRandomOptions);
var csrfRandomNumber;

bot.dialog('/', function(session){
    //Commenting this part .. need to test this 
    //connector.getUserToken(session.message.address, process.env.CONNECTION, undefined, function(err, result) {
    //    if(result) {
    //        session.send("You are already signed in (SDK)");
    //    }
    //    else {
    //        console.log(err);
    //        if(!session.userData.accessKey) {
    //            botBuilder.OAuthCard.create(connector, session, process.env.CONNECTION, "Please sign in", function(createSignErr, signInMessage) {
    //                if(signInMessage) {
    //                    session.send(signInMessage);
    //                    session.userData.accessKey = 1;
    //                }
    //                else {
    //                    session.send("Issue with your signin: %s", createSignErr);
    //                }
    //            })
    //        }
    //    }
    //});
    tenantId = teams.TeamsMessage.getTenantId(session.message);
    //tempAddress = session.message.address;
    tempAddress = session.message.address.conversation.id;
    if(session.userData.accessKey) {
        session.send("Hi %s, you are already logged in", session.message.user.name);
    }
    else if( userKey[session.message.user.name]) {
        session.userData.accessKey = userKey[session.message.user.name];
        session.send("Hi %s, you are already logged in", session.message.user.name);
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
   res.redirect(loginURL, next);
});

server.use(restify.plugins.queryParser());
server.get("/verified", function(req, res, next){
   if(parseInt(req.query.state) !== csrfRandomNumber) res.send(401, "CSRF error");
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
           var json = JSON.parse(body);
           var decoded = jwtDecoder(json["id_token"]);
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
               }
           });
           userKey[decoded.name] = "authenticated";
           var botApiKeyOptions = {
               headers: {
                            "Content-type":  "application/x-www-form-urlencoded",
               },
               body: "grant_type=client_credentials&client_id=" + encodeURIComponent(process.env.MICROSOFT_APP_ID) +
                     "&client_secret=" + encodeURIComponent(process.env.MICROSOFT_APP_PASSWORD) + "&scope=https%3A%2F%2Fapi.botframework.com%2F.default"
           }
           request.get("https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token", botApiKeyOptions, function(error, response, body){
                if(error) console.log("Error while accessing Api Token for bot" + error);
                var accessKeyJson = JSON.parse(body);
                var botPostHeaders = {
                    headers: {
                        "Authorization" : "Bearer " + accessKeyJson["access_token"],
                        "Content-Type"  : "application/json"
                    },
                    json : {
                        "type": "message",
                        "from": {
                            "id": process.env.MICROSOFT_APP_ID,
                            "name": "Trackerbot"
                        },
                        "text": "You have been authenticated ( web)"
                        }
                }
                request.post("https://smba.trafficmanager.net/amer/v3/conversations/" + encodeURIComponent(tempAddress) + "/activities", botPostHeaders, function(error, response, body){
                    if(error) console.log(error);
                    console.log(body);
                });
           })
           res.send(200, "Successfully authenticated");
           next();
       });
   }
});

server.get("/", function(req, response){ response.send(200,"Your app is up and running.")})

server.get("/terms", function(req, response){ response.send(200, "Sample terms of Usage page")})

server.get("/privacy", function(req, response){ response.send(200, "Privacy policy")});
