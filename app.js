var restify          = require("restify");
var botBuilder       = require("botbuilder");
var request          = require("request");
var jwtDecoder       = require("jwt-decode");
var jsonwebtoken     = require("jsonwebtoken");
var jwkToPem         = require("jwk-to-pem");
var redis            = require("redis");
var crypto           = require('crypto'), algorithm = 'aes-128-cbc', password = process.env.CRYPTO_PASS;

function encrypt(text){
  var cipher = crypto.createCipher(algorithm,password)
  var crypted = cipher.update(text,'utf8','hex')
  crypted += cipher.final('hex');
  return crypted;
}

function decrypt(text){
  var decipher = crypto.createDecipher(algorithm,password)
  var dec = decipher.update(text,'hex','utf8')
  dec += decipher.final('utf8');
  return dec;
}

var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3980, function () {
    console.log('%s listening to %s', server.name, server.url); 
 });

var redisHost = process.env.REDIS_URL;
var client = redis.createClient({url: redisHost, password: process.env.REDIS_PASS});
client.on('connect', function(){
    console.log('Redis client connected');
});
client.on('error', function(err){
    console.log('Error: %s in connecting Redis client', err);
});

function saveData(context, data, callback){
    client.hset('userData', encrypt(JSON.stringify(context)), encrypt(JSON.stringify(data)), function(err, res){
        callback(err);
    });
}

function getData(context, callback){
    client.hget('userData', encrypt(JSON.stringify(context)), function(err, result){
        if(err) callback(err,{});
        else if(!result) callback({userData : undefined});
        else callback(JSON.parse(decrypt(result)));
    })
}

server.use(restify.plugins.bodyParser());
server.post('/api/messages',function(req, res, next){
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
                                getData({"persistConversationData" : "true", "persistUserData" : "true", "userId" : activityJson.from.id}, function(err, data){
                                    if(err) console.log("Error at getData: " +  err);
                                    console.log("getData SDK: %s",data.userData);
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
                                                    + "," + activityJson.from.id + ")"
                                        }
                                    }
                                    if(data.userData) {
                                        var tokenJson =  JSON.parse(new Buffer(data.userData, 'base64').toString('ascii'));
                                        console.log("Access token: %s",tokenJson.access_token);
                                        var graphGetOptions = {
                                            headers : {
                                                "Content-Type" : "application/json",
                                                "Authorization": tokenJson.access_token
                                            }
                                        }
                                        request.get("https://graph.microsoft.com/v1.0/me", graphGetOptions, function(graphErr, graphRes, graphBody){
                                            if(graphErr) console.log(graphErr);
                                            if(graphRes.statusCode !== 200) {
                                                console.log("Refreshing token for user: " + decoded["unique_name"] );
                                                var refreshOptions = {
                                                    host: "https://login.microsoftonline.com",
                                                    headers: {
                                                        "Content-type":  "application/x-www-form-urlencoded",
                                                    },
                                                    body:   "grant_type=refresh_token&client_id=" + process.env.APP_CLIENT_ID +
                                                            "&refresh_token=" + tokenJson.refresh_token +
                                                            "&resource=" + encodeURIComponent("https://graph.microsoft.com") +
                                                            "&client_secret=" + encodeURIComponent(process.env.APP_KEY)
                                                }
                                                request.post("https://login.microsoftonline.com/" + process.env.TenantId + "/oauth2/token", refreshOptions, function(refreshError, refreshResponse, refreshBody){
                                                    console.log(refreshError);
                                                   // console.log(refreshBody);
                                                    var refreshJson = JSON.parse(refreshBody);
                                                    graphGetOptions.headers.Authorization = refreshJson.access_token;
                                                    if(refreshJson.error) console.log(refreshJson.error);
                                                    else {
                                                        request.get("https://graph.microsoft.com/v1.0/me", graphGetOptions, function(refreshGetError, refreshGetResponse, refreshGetBody){
                                                            if(refreshGetError)console.log(refreshGetError);
                                                            else {
                                                                console.log(refreshGetBody);
                                                                botPostHeaders.json.text = "You are logged in"
                                                                request.post("https://smba.trafficmanager.net/amer/v3/conversations/" + encodeURIComponent(activityJson.conversation.id) + "/activities", 
                                                                              botPostHeaders, function(error, response, body){
                                                                                    if(error) console.log(error);
                                                                                    console.log(body);
                                                                });
                                                                var oauthAccessToken = new Buffer(JSON.stringify(refreshJson)).toString('base64');
                                                                saveData({"persistConversationData" : "true", "persistUserData" : "true", "userId" : activityJson.from.id}, {"userData" : oauthAccessToken} , function(err){
                                                                    if(err) console.log("saveData Err:" + err);
                                                                });
                                                            }
                                                        });
                                                    }
                                                });
                                            }
                                            else {
                                                console.log(graphBody);
                                                botPostHeaders.json.text = "You are logged in"
                                                request.post("https://smba.trafficmanager.net/amer/v3/conversations/" + encodeURIComponent(activityJson.conversation.id) + "/activities", 
                                                              botPostHeaders, function(error, response, body){
                                                                    if(error) console.log(error);
                                                                    console.log(body);
                                                });
                                            }
                                        });
                                    }
                                    else {
                                        request.post("https://smba.trafficmanager.net/amer/v3/conversations/" + encodeURIComponent(activityJson.conversation.id) + "/activities", 
                                                      botPostHeaders, function(error, response, body){
                                                         if(error) console.log(error);
                                                        console.log(body);
                                        });
                                    }
                                });
                            })
                        })
                    });
                });
            }
        })
    }
});

//server.post("/api/messages", connector.listen());

server.use(restify.plugins.queryParser());
server.get("/verified", function(req, res, next){
   if(!req.query.state.match(process.env.csrfToken)) res.send(401, "CSRF error");
   else {
       var tempArr = req.query.state.split(",");
       console.log("tempArr:" + req.query.state);
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
           if(error) console.log(error);
           var json = JSON.parse(body);
           console.log(json);
           var decoded = jwtDecoder(json["id_token"]);
           var getOptions = {
               headers: {
                   "Content-Type"  : "application/json",
                   "Authorization" : json["access_token"]
               }
           }
           var oauthAccessToken = new Buffer(JSON.stringify(json)).toString('base64');
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
                       var refreshJson = JSON.parse(refreshBody);
                       getOptions.headers.Authorization = refreshJson.access_token;
                       if(refreshJson.error) console.log(refreshJson.error);
                       else {
                           request.get("https://graph.microsoft.com/v1.0/me", getOptions, function(refreshGetError, refreshGetResponse, refreshGetBody){
                               console.log(refreshGetError);
                           });
                       }
                   });
               }
               else {
                   console.log(getError);
               }
           });
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
                console.log("conv, user:" + conversationId + ","+ userId);
                request.post("https://smba.trafficmanager.net/amer/v3/conversations/" + conversationId + "/activities", 
                             botPostHeaders, function(error, response, body){
                    if(error) console.log(error);
                    console.log("conv update: " + body);
                });
                botPostHeaders.json = { "data" : oauthAccessToken, "eTag" : "test"}
                console.log("oauthToken: %s",oauthAccessToken);
                saveData({"persistConversationData" : "true", "persistUserData" : "true", "userId" : userId}, {"userData" : oauthAccessToken} , function(err){
                    if(err) console.log("saveData Err:" + err);
                    getData({"persistConversationData" : "true", "persistUserData" : "true", "userId" : userId}, function(err, data){
                        if(err) console.log(err);
                        console.log("User State Data: ", JSON.stringify(data));
                    });
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
