var restify          = require("restify");
var botBuilder       = require("botbuilder");
var teams            = require("botbuilder-teams");
var request          = require("request");
var jwtDecoder       = require("jwt-decode");
var rn               = require('random-number');
var jsonwebtoken     = require("jsonwebtoken");
var jwkToPem         = require("jwk-to-pem");


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
    //console.log(req.headers);
    if(req.headers.authorization.includes("Bearer "))
        res.send(403,"Forbidden");
    var activityJson = req.body;
    var loginURL = "https://login.microsoftonline.com/" + activityJson.channelData.tenant.id + "/oauth2/authorize?client_id=" + process.env.APP_CLIENT_ID +
                   "&response_type=code&redirect_uri=" + encodeURIComponent( "https://" + req.headers.host + "/verified") +
                   "&response_mode=query&resource=" + encodeURIComponent("https://graph.microsoft.com")+ "&state=";

    var bearerToken = req.headers.authorization.replace("Bearer ","");
    var arr = bearerToken.split('.');
    arr.pop();
    var jwtPayload = JSON.parse(new Buffer(arr.pop(), 'base64').toString('ascii'));
    var jwtHeader = JSON.parse(new Buffer(arr.pop(), 'base64').toString('ascii'));
    if(jwtPayload.aud !== process.env.MICROSOFT_APP_ID) {
        console.log("aud not matched");
        //res.writeHead(403, {'Content-Type' : 'text/html'});
        //res.write("Forbidden");
        //res.end();
    }
    else {
        request.get("https://login.botframework.com/v1/.well-known/openidconfiguration", function(getReq, getRes, next){
            var getJson = JSON.parse(getRes.body);
            if(jwtPayload.iss !== getJson.issuer || jwtHeader.alg != getJson.id_token_signing_alg_values_supported || jwtPayload.exp < Math.round((new Date).getTime() / 1000) || 
                jwtPayload.serviceurl !== activityJson.serviceUrl) {
                    console.log("conf not matched");
                    console.log(jwtPayload.iss !== getJson.issuer);
                    console.log(jwtHeader.alg != getJson.id_token_signing_alg_values_supported);
                    console.log(jwtPayload.serviceurl !== activityJson.serviceUrl);
                    console.log(jwtPayload.exp < (new Date).getTime());
                    //res.write("Forbidden");
                    //res.end();
            }
            else {
                console.log(getJson.id_token_signing_alg_values_supported);
                console.log(getJson.jwks_uri);
                request.get(getJson.jwks_uri, function(keyReq,keyRes,keyNext){
                    var keysArray = JSON.parse(keyRes.body).keys;
                    keysArray.forEach(function(key){
                        if(key.kid !== jwtHeader.kid) return;
                        console.log("valid Key: " + jwtHeader.kid);
                        console.log("listed Key: " + key.kid);
                        console.log("alg:" + getJson.id_token_signing_alg_values_supported[0] + typeof getJson.id_token_signing_alg_values_supported[0]);
                        jsonwebtoken.verify(bearerToken, jwkToPem(key), {"algorithm" : "RS256", "expiresIn" : jwtPayload.exp} , function(err, decoded){
                            if(err) console.log(err);
                            console.log(decoded);
                            console.log(activityJson);
                            // TODO: Use Session object not botstate api call. 
                            //var session = new botBuilder.Session({"connector" : connector, "dialogId" : activityJson.conversation.id,
                            //                                       "library" : new botBuilder.library()
                            //});
                            //session.userData.test = "test " + activityJson.id;
                            //console.log(session.userData);
                            //session.send("Hi");
                            var botApiKeyOptions = {
                                headers: {
                                             "Content-type":  "application/x-www-form-urlencoded",
                                },
                                body: "grant_type=client_credentials&client_id=" + encodeURIComponent(process.env.MICROSOFT_APP_ID) +
                                      "&client_secret=" + encodeURIComponent(process.env.MICROSOFT_APP_PASSWORD) + "&scope=https%3A%2F%2Fapi.botframework.com%2F.default"
                            }
                            request.get("https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token", botApiKeyOptions, function(error, response, body){
                                if(error) console.log("Error while accessing Api Token for bot" + error);
                                console.log(body);
                                console.log(activityJson.conversation.id + activityJson.from.id);
                                var accessKeyJson = JSON.parse(body);
                                var botGetHeaders = {
                                    headers: {
                                        "Authorization" : "Bearer " + accessKeyJson["access_token"],
                                        "Content-Type"  : "application/json"
                                    },
                                }
                                request.get("https://smba.trafficmanager.net/amer/v3/botstate/" + encodeURIComponent(activityJson.channelId) + "/users/" + encodeURIComponent(activityJson.from.id),
                                              botGetHeaders, function(error, response, body){
                                                if(error) console.log(error);
                                                console.log("Bot state get: " + body);
                                                var botPostHeaders = {
                                                    headers: {
                                                        "Authorization" : "Bearer " + accessKeyJson["access_token"],
                                                        "Content-Type"  : "application/json"
                                                    },
                                                    json: {
                                                        "type": "message",
                                                        "from": {
                                                            "id": process.env.MICROSOFT_APP_ID,
                                                            "name": "Trackerbot"
                                                        },
                                                        "text": "Please [login](" + loginURL + process.env.csrfToken + "," + activityJson.conversation.id
                                                                + "," + activityJson.from.Id + ")"
                                                    }
                                                }
                                                if(body) {
                                                    botPostHeaders.json.text = "You are logged in"
                                                }
                                                request.post("https://smba.trafficmanager.net/amer/v3/conversations/" + encodeURIComponent(activityJson.conversation.id) + "/activities", 
                                                              botPostHeaders, function(error, response, body){
                                                                 if(error) console.log(error);
                                                                console.log(body);
                                                });
                                            });
                            })
                            var graphGetOptions = {
                                headers : {
                                    "Content-Type" : "application/json",
                                    "Authorization": req.headers.authorization
                                }
                            }
                           // request.get("https://graph.microsoft.com/v1.0/me", graphGetOptions, function(graphErr, graphRes, graphBody){
                           //     if(graphErr) console.log(graphErr);
                           //     console.log(graphBody);
                           // });
                        })
                    });
                });
            }
        })
    }
});

//server.post("/api/messages", connector.listen());

var userKey = {};

var tempAddress;

var csrfRandomOptions = { min: 0, max: 100000, integer: true};
var generateRandom = rn.generator(csrfRandomOptions);
var csrfRandomNumber;

bot.dialog('/', function(session){
    console.log(session);
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
    var loginURL = "https://login.microsoftonline.com/" + process.env.TenantId + "/oauth2/authorize?client_id=" + process.env.APP_CLIENT_ID +
                   "&response_type=code&redirect_uri=" + encodeURIComponent( "https://" + req.headers.host + "/verified") +
                   "&response_mode=query&resource=" + encodeURIComponent("https://graph.microsoft.com")+ "&state=" + csrfRandomNumber;
   res.redirect(loginURL, next);
});

server.use(restify.plugins.queryParser());
server.get("/verified", function(req, res, next){
   //if(parseInt(req.query.state) !== csrfRandomNumber) res.send(401, "CSRF error");
   if(!req.query.state.match(process.env.csrfToken)) res.send(401, "CSRF error");
   else {
       var tempArr = req.query.state.split(",");
       var userId = tempArr.pop();
       var conversationId = tempArr.pop();
       var authURLOptions = {
           host: "https://login.microsoftonline.com/",
           path: process.env.TenantId + "/oauth2/token",
           headers: {
               "Content-type":  "application/x-www-form-urlencoded",
           },
           body: "grant_type=authorization_code&client_id=" + process.env.APP_CLIENT_ID + "&code="+ req.query.code + "&redirect_uri=" + 
               encodeURIComponent("https://" + req.headers.host + "/verified") + "&resource=" + encodeURIComponent("https://graph.microsoft.com") +
               "&client_secret=" + encodeURIComponent(process.env.APP_KEY)
       }
       request.post("https://login.microsoftonline.com/" + process.env.TenantId + "/oauth2/token", authURLOptions, function(error, response, body){
           console.log(error);
           var json = JSON.parse(body);
           console.log(json);
           var decoded = jwtDecoder(json["id_token"]);
           var getOptions = {
               headers: {
                   "Content-Type"  : "application/json",
                   "Authorization" : json["access_token"]
               }
           }
           var oauthAccessToken = Buffer.from(json.toString()).toString("base64");
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
                   request.post("https://login.microsoftonline.com/" + process.env.TenantId + "/oauth2/token", refreshOptions, function(refreshError, refreshResponse, refreshBody){
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
                console.log(conversationId + ","+ userId);
                request.post("https://smba.trafficmanager.net/amer/v3/conversations/" + conversationId + "/activities", 
                             botPostHeaders, function(error, response, body){
                    if(error) console.log(error);
                    console.log("conv update: " + body);
                });
                botPostHeaders.json = { "data" : oauthAccessToken, "eTag" : "test"}
                request.post("https://smba.trafficmanager.net/amer/v3/botstate/msteams/users/" + userId,
                              botPostHeaders, function(error, response, body){
                    if(error) console.log(error);
                    console.log("state data:" + body);
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
