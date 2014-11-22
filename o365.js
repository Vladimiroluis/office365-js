/// Microsoft.Office365.ClientLib.JS.1.0.22
/// patched by masataka takeuchi for nodejs support
var window = require("./lib/node-window"), XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var Microsoft;
(function (Microsoft) {
    (function (Utility) {
        (function (EncodingHelpers) {
            function getKeyExpression(entityKeys) {
                var entityInstanceKey = '(';

                if (entityKeys.length == 1) {
                    entityInstanceKey += formatLiteral(entityKeys[0]);
                } else {
                    var addComma = false;
                    for (var i = 0; i < entityKeys.length; i++) {
                        if (addComma) {
                            entityInstanceKey += ',';
                        } else {
                            addComma = true;
                        }

                        entityInstanceKey += entityKeys[i].name + '=' + formatLiteral(entityKeys[i]);
                    }
                }

                entityInstanceKey += ')';

                return entityInstanceKey;
            }
            EncodingHelpers.getKeyExpression = getKeyExpression;

            function formatLiteral(literal) {
                /// <summary>Formats a value according to Uri literal format</summary>
                /// <param name="value">Value to be formatted.</param>
                /// <param name="type">Edm type of the value</param>
                /// <returns type="string">Value after formatting</returns>
                var result = "" + formatRowLiteral(literal.value, literal.type);

                result = encodeURIComponent(result.replace("'", "''"));

                switch ((literal.type)) {
                    case "Edm.Binary":
                        return "X'" + result + "'";
                    case "Edm.DateTime":
                        return "datetime" + "'" + result + "'";
                    case "Edm.DateTimeOffset":
                        return "datetimeoffset" + "'" + result + "'";
                    case "Edm.Decimal":
                        return result + "M";
                    case "Edm.Guid":
                        return "guid" + "'" + result + "'";
                    case "Edm.Int64":
                        return result + "L";
                    case "Edm.Float":
                        return result + "f";
                    case "Edm.Double":
                        return result + "D";
                    case "Edm.Geography":
                        return "geography" + "'" + result + "'";
                    case "Edm.Geometry":
                        return "geometry" + "'" + result + "'";
                    case "Edm.Time":
                        return "time" + "'" + result + "'";
                    case "Edm.String":
                        return "'" + result + "'";
                    default:
                        return result;
                }
            }
            EncodingHelpers.formatLiteral = formatLiteral;

            function formatRowLiteral(value, type) {
                switch (type) {
                    case "Edm.Binary":
                        return Microsoft.Utility.decodeBase64AsHexString(value);
                    default:
                        return value;
                }
            }
        })(Utility.EncodingHelpers || (Utility.EncodingHelpers = {}));
        var EncodingHelpers = Utility.EncodingHelpers;

        function findProperties(o) {
            var aPropertiesAndMethods = [];

            do {
                aPropertiesAndMethods = aPropertiesAndMethods.concat(Object.getOwnPropertyNames(o));
            } while(o = Object.getPrototypeOf(o));

            for (var a = 0; a < aPropertiesAndMethods.length; ++a) {
                for (var b = a + 1; b < aPropertiesAndMethods.length; ++b) {
                    if (aPropertiesAndMethods[a] === aPropertiesAndMethods[b]) {
                        aPropertiesAndMethods.splice(a--, 1);
                    }
                }
            }

            return aPropertiesAndMethods;
        }
        Utility.findProperties = findProperties;

        function decodeBase64AsHexString(base64) {
            var decoded = decodeBase64(base64), hexValue = "", hexValues = "0123456789ABCDEF";

            for (var j = 0; j < decoded.length; j++) {
                var byte = decoded[j];
                hexValue += hexValues[byte >> 4];
                hexValue += hexValues[byte & 0x0F];
            }

            return hexValue;
        }
        Utility.decodeBase64AsHexString = decodeBase64AsHexString;

        function decodeBase64(base64) {
            var decoded = [];

            if (window.atob !== undefined) {
                var binaryStr = window.atob(base64);
                for (var i = 0; i < binaryStr.length; i++) {
                    decoded.push(binaryStr.charCodeAt(i));
                }
                return decoded;
            }

            for (var index = 0; index < base64.length; index += 4) {
                var sextet1 = getBase64Sextet(base64[index]);
                var sextet2 = getBase64Sextet(base64[index + 1]);
                var sextet3 = (index + 2 < base64.length) ? getBase64Sextet(base64[index + 2]) : null;
                var sextet4 = (index + 3 < base64.length) ? getBase64Sextet(base64[index + 3]) : null;
                decoded.push((sextet1 << 2) | (sextet2 >> 4));
                if (sextet3)
                    decoded.push(((sextet2 & 0xF) << 4) | (sextet3 >> 2));
                if (sextet4)
                    decoded.push(((sextet3 & 0x3) << 6) | sextet4);
            }

            return decoded;
        }
        Utility.decodeBase64 = decodeBase64;

        function decodeBase64AsString(base64) {
            var decoded = decodeBase64(base64), decoded_string;

            decoded.forEach(function (value, index, decoded_access_token) {
                if (!decoded_string) {
                    decoded_string = String.fromCharCode(value);
                } else {
                    decoded_string += String.fromCharCode(value);
                }
            });

            return decoded_string;
        }
        Utility.decodeBase64AsString = decodeBase64AsString;

        function getBase64Sextet(character) {
            var code = character.charCodeAt(0);

            if (code >= 65 && code <= 90)
                return code - 65;

            if (code >= 97 && code <= 122)
                return code - 71;

            if (code >= 48 && code <= 57)
                return code + 4;

            if (character === "+")
                return 62;

            if (character === "/")
                return 63;

            return null;
        }

        var Exception = (function () {
            function Exception(message, innerException) {
                this._message = message;
                if (innerException) {
                    this._innerException = innerException;
                }
            }
            Object.defineProperty(Exception.prototype, "message", {
                get: function () {
                    return this._message;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Exception.prototype, "innerException", {
                get: function () {
                    return this._innerException;
                },
                enumerable: true,
                configurable: true
            });
            return Exception;
        })();
        Utility.Exception = Exception;

        var HttpException = (function (_super) {
            __extends(HttpException, _super);
            function HttpException(XHR, innerException) {
                _super.call(this, XHR.statusText, innerException);
                this.getHeaders = this.getHeadersFn(XHR);
            }
            HttpException.prototype.getHeadersFn = function (xhr) {
                return function (headerName) {
                    if (headerName && headerName.length > 0) {
                        return xhr.getResponseHeader(headerName);
                    } else {
                        return xhr.getAllResponseHeaders();
                    }
                    ;
                };
            };

            Object.defineProperty(HttpException.prototype, "xhr", {
                get: function () {
                    return this._xhr;
                },
                enumerable: true,
                configurable: true
            });
            return HttpException;
        })(Exception);
        Utility.HttpException = HttpException;

        var DeferredState;
        (function (DeferredState) {
            DeferredState[DeferredState["UNFULFILLED"] = 0] = "UNFULFILLED";
            DeferredState[DeferredState["RESOLVED"] = 1] = "RESOLVED";
            DeferredState[DeferredState["REJECTED"] = 2] = "REJECTED";
        })(DeferredState || (DeferredState = {}));

        var Deferred = (function () {
            function Deferred() {
                this._fulfilled = function (value) {
                };
                this._rejected = function (reason) {
                };
                this._progress = function (progress) {
                };
                this._state = 0 /* UNFULFILLED */;
            }
            Deferred.prototype.then = function (onFulfilled, onRejected, onProgress) {
                switch (this._state) {
                    case 0 /* UNFULFILLED */:
                        if (onFulfilled && typeof onFulfilled === 'function') {
                            var fulfilled = this._fulfilled;
                            this._fulfilled = function (value) {
                                fulfilled(value);
                                onFulfilled(value);
                            };
                        }
                        if (onRejected && typeof onRejected === 'function') {
                            var rejected = this._rejected;
                            this._rejected = function (reason) {
                                rejected(reason);
                                onRejected(reason);
                            };
                        }
                        if (onProgress && typeof onProgress === 'function') {
                            var progress = this._progress;
                            this._progress = function (progress) {
                                progress(progress);
                                onProgress(progress);
                            };
                        }
                        break;
                    case 1 /* RESOLVED */:
                        if (onFulfilled && typeof onFulfilled === 'function') {
                            onFulfilled(this._value);
                        }
                        break;
                    case 2 /* REJECTED */:
                        if (onRejected && typeof onRejected === 'function') {
                            onRejected(this._reason);
                        }
                        break;
                }

                return this;
            };

            Deferred.prototype.detach = function () {
                this._fulfilled = function (value) {
                };
                this._rejected = function (reason) {
                };
                this._progress = function (progress) {
                };
            };

            Deferred.prototype.resolve = function (value) {
                if (this._state != 0 /* UNFULFILLED */) {
                    throw new Microsoft.Utility.Exception("Invalid deferred state = " + this._state);
                }
                this._value = value;
                var fulfilled = this._fulfilled;
                this.detach();
                this._state = 1 /* RESOLVED */;
                fulfilled(value);
            };

            Deferred.prototype.reject = function (reason) {
                if (this._state != 0 /* UNFULFILLED */) {
                    throw new Microsoft.Utility.Exception("Invalid deferred state = " + this._state);
                }
                this._reason = reason;
                var rejected = this._rejected;
                this.detach();
                this._state = 2 /* REJECTED */;
                rejected(reason);
            };

            Deferred.prototype.notify = function (progress) {
                if (this._state != 0 /* UNFULFILLED */) {
                    throw new Microsoft.Utility.Exception("Invalid deferred state = " + this._state);
                }
                this._progress(progress);
            };
            return Deferred;
        })();
        Utility.Deferred = Deferred;

        (function (HttpHelpers) {
            var Request = (function () {
                function Request(requestUri, method, data) {
                    this.requestUri = requestUri;
                    this.method = method;
                    this.data = data;
                    this.headers = {};
                    this.disableCache = false;
                }
                return Request;
            })();
            HttpHelpers.Request = Request;

            var AuthenticatedHttp = (function () {
                function AuthenticatedHttp(getAccessTokenFn) {
                    this._disableCache = false;
                    this._noCache = Date.now();
                    this._accept = 'application/json;q=0.9, */*;q=0.1';
                    this._contentType = 'application/json';
                    this._getAccessTokenFn = getAccessTokenFn;
                }
                Object.defineProperty(AuthenticatedHttp.prototype, "disableCache", {
                    get: function () {
                        return this._disableCache;
                    },
                    set: function (value) {
                        this._disableCache = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                Object.defineProperty(AuthenticatedHttp.prototype, "accept", {
                    get: function () {
                        return this._accept;
                    },
                    set: function (value) {
                        this._accept = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                Object.defineProperty(AuthenticatedHttp.prototype, "contentType", {
                    get: function () {
                        return this._contentType;
                    },
                    set: function (value) {
                        this._contentType = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                AuthenticatedHttp.prototype.ajax = function (request) {
                    var deferred = new Microsoft.Utility.Deferred();

                    var xhr = new XMLHttpRequest(); xhr.responseType = request.responseType;

                    if (!request.method) {
                        request.method = 'GET';
                    }

                    xhr.open(request.method.toUpperCase(), request.requestUri, true);

                    if (request.headers) {
                        for (name in request.headers) {
                            xhr.setRequestHeader(name, request.headers[name]);
                        }
                    }

                    xhr.onreadystatechange = function (e) {
                        if (xhr.readyState == 4) {
                            if (xhr.status >= 200 && xhr.status < 300 || xhr.status === 304) {
                                deferred.resolve(xhr.response);
                            } else {
                                deferred.reject(xhr);
                            }
                        } else {
                            deferred.notify(xhr.readyState);
                        }
                    };

                    if (request.data) {
                        if ((typeof request.data === 'string') || (Buffer.isBuffer(request.data))) {
                            xhr.send(request.data);
                        } else {
                            xhr.send(JSON.stringify(request.data));
                        }
                    } else {
                        xhr.send();
                    }

                    return deferred;
                };

                AuthenticatedHttp.prototype.getUrl = function (url) {
                    return this.request(new Request(url));
                };

                AuthenticatedHttp.prototype.postUrl = function (url, data) {
                    return this.request(new Request(url, 'POST', data));
                };

                AuthenticatedHttp.prototype.deleteUrl = function (url) {
                    return this.request(new Request(url, 'DELETE'));
                };

                AuthenticatedHttp.prototype.patchUrl = function (url, data) {
                    return this.request(new Request(url, 'PATCH', data));
                };

                AuthenticatedHttp.prototype.request = function (request) {
                    var _this = this;
                    var deferred;

                    this.augmentRequest(request);

                    if (this._getAccessTokenFn) {
                        deferred = new Microsoft.Utility.Deferred();

                        this._getAccessTokenFn().then((function (token) {
                            request.headers["Authorization"] = 'Bearer ' + token;
                            _this.ajax(request).then(deferred.resolve, deferred.reject);
                        }).bind(this), deferred.reject);
                    } else {
                        deferred = this.ajax(request);
                    }

                    return deferred;
                };

                AuthenticatedHttp.prototype.augmentRequest = function (request) {
                    if (!request.headers) {
                        request.headers = {};
                    }

                    if (!request.headers['Accept']) {
                        request.headers['Accept'] = this._accept;
                    }

                    if (!request.headers['Content-Type']) {
                        request.headers['Content-Type'] = this._contentType;
                    }

                    if (request.disableCache || this._disableCache) {
                        request.requestUri += (request.requestUri.indexOf('?') >= 0 ? '&' : '?') + '_=' + this._noCache++;
                    }
                };
                return AuthenticatedHttp;
            })();
            HttpHelpers.AuthenticatedHttp = AuthenticatedHttp;
        })(Utility.HttpHelpers || (Utility.HttpHelpers = {}));
        var HttpHelpers = Utility.HttpHelpers;
    })(Microsoft.Utility || (Microsoft.Utility = {}));
    var Utility = Microsoft.Utility;
})(Microsoft || (Microsoft = {}));
//# sourceMappingURL=utility.js.map
var O365Auth;
(function (O365Auth) {
    var Token = (function () {
        function Token(idToken, context, resourceId, clientId, redirectUri) {
            var encoded_idToken = idToken.split('.')[1].replace('-', '+').replace('_', '/'), decoded_idToken = Microsoft.Utility.decodeBase64AsString(encoded_idToken);

            this._idToken = JSON.parse(decoded_idToken);
            this._context = context;
            this._clientId = clientId || O365Auth.Settings.clientId;
            this._redirectUri = redirectUri || O365Auth.Settings.redirectUri;
            this._resourceId = resourceId;
        }
        Token.prototype.getDeferred = function () {
            if (O365Auth.deferred) {
                return O365Auth.deferred();
            }

            return new Microsoft.Utility.Deferred();
        };

        Token.prototype.getAccessTokenFn = function (resourceId) {
            return function () {
                return this.getAccessToken(resourceId || this._resourceId);
            }.bind(this);
        };

        Token.prototype.getAccessToken = function (resourceId) {
            return this._context.getAccessToken(resourceId, null, this._clientId, this._redirectUri);
        };

        Object.defineProperty(Token.prototype, "audience", {
            get: function () {
                return this._idToken.aud;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Token.prototype, "familyName", {
            get: function () {
                return this._idToken.family_name;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Token.prototype, "givenName", {
            get: function () {
                return this._idToken.given_name;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Token.prototype, "identityProvider", {
            get: function () {
                return this._idToken.iss;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Token.prototype, "objectId", {
            get: function () {
                return this._idToken.oid;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Token.prototype, "tenantId", {
            get: function () {
                return this._idToken.tid;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Token.prototype, "uniqueName", {
            get: function () {
                return this._idToken.unique_name;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Token.prototype, "userPrincipalName", {
            get: function () {
                return this._idToken.upn;
            },
            enumerable: true,
            configurable: true
        });
        return Token;
    })();
    O365Auth.Token = Token;

    var CacheManager = (function () {
        function CacheManager() {
            this._client_id_key = 'client_id';
            this._access_tokens_key = 'access_tokens';
            this._refresh_token_key = 'refresh_token';
            this._idToken_key = 'id_token';
            try  {
                var cache_entry = window.localStorage.getItem(this._client_id_key);

                if (cache_entry === null || cache_entry.length === 0) {
                    cache_entry = '{}';
                }

                this._cache_entry = JSON.parse(cache_entry);

                if (!this._cache_entry || typeof this._cache_entry === 'string') {
                    this._cache_entry = {};
                }
            } catch (e) {
                this._cache_entry = {};
            }
        }
        CacheManager.prototype.save = function () {
            try  {
                window.localStorage.setItem(this._client_id_key, JSON.stringify(this._cache_entry));
            } catch (e) {
            }
        };

        CacheManager.prototype.clearAll = function () {
            this._cache_entry = {};
            this.save();
        };

        CacheManager.prototype.clear = function (client_id) {
            this._cache_entry[client_id] = undefined;
            this.save();
        };

        CacheManager.prototype.getAccessToken = function (client_id, resource_id) {
            this._cache_entry[client_id] = this._cache_entry[client_id] || {};
            this._cache_entry[client_id][this._access_tokens_key] = this._cache_entry[client_id][this._access_tokens_key] || {};
            if (this._cache_entry[client_id][this._access_tokens_key][resource_id] && typeof this._cache_entry[client_id][this._access_tokens_key][resource_id].expires_in === 'string') {
                this._cache_entry[client_id][this._access_tokens_key][resource_id].expires_in = new Date(this._cache_entry[client_id][this._access_tokens_key][resource_id].expires_in);
            }
            return this._cache_entry[client_id][this._access_tokens_key][resource_id];
        };

        CacheManager.prototype.getRefreshToken = function (client_id) {
            this._cache_entry[client_id] = this._cache_entry[client_id] || {};
            return this._cache_entry[client_id][this._refresh_token_key];
        };

        CacheManager.prototype.getIdToken = function (client_id) {
            this._cache_entry[client_id] = this._cache_entry[client_id] || {};
            return this._cache_entry[client_id][this._idToken_key];
        };

        CacheManager.prototype.setAccessToken = function (client_id, resource_id, access_token) {
            this._cache_entry[client_id] = this._cache_entry[client_id] || {};
            this._cache_entry[client_id][this._access_tokens_key] = this._cache_entry[client_id][this._access_tokens_key] || {};
            this._cache_entry[client_id][this._access_tokens_key][resource_id] = access_token;
            this.save();
        };

        CacheManager.prototype.setRefreshToken = function (client_id, refresh_token) {
            this._cache_entry[client_id] = this._cache_entry[client_id] || {};
            this._cache_entry[client_id][this._refresh_token_key] = refresh_token;
            this.save();
        };

        CacheManager.prototype.setIdToken = function (client_id, id_token) {
            this._cache_entry[client_id] = this._cache_entry[client_id] || {};
            this._cache_entry[client_id][this._idToken_key] = id_token;
            this.save();
        };
        return CacheManager;
    })();

    O365Auth.deferred;

    var Context = (function () {
        function Context(authUri, redirectUri) {
            this._redirectUri = 'http://localhost/';
            this._cacheManager = new CacheManager();
            if (!authUri) {
                if (O365Auth.Settings.authUri) {
                    this._authUri = O365Auth.Settings.authUri;
                } else {
                    throw new Microsoft.Utility.Exception('No authUri provided nor found in O365Auth.authUri');
                }
            } else {
                this._authUri = authUri;
            }
            if (this._authUri.charAt(this._authUri.length - 1) !== '/') {
                this._authUri += '/';
            }
            if (!redirectUri) {
                if (O365Auth.Settings.redirectUri) {
                    this._redirectUri = O365Auth.Settings.redirectUri;
                }
            } else {
                this._redirectUri = redirectUri;
            }
        }
        Context.prototype.getDeferred = function () {
            if (O365Auth.deferred) {
                return O365Auth.deferred();
            }

            return new Microsoft.Utility.Deferred();
        };

        Context.prototype.ajax = function (url, data, verb) {
            var deferred = new Microsoft.Utility.Deferred(), xhr = new XMLHttpRequest();

            if (!verb) {
                verb = 'GET';
            }

            xhr.open(verb.toUpperCase(), url, true);

            xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded; charset=UTF-8');
            xhr.setRequestHeader('Accept', '*/*');

            xhr.onreadystatechange = function (e) {
                if (xhr.readyState == 4) {
                    if (xhr.status >= 200 && xhr.status < 300 || xhr.status === 304) {
                        deferred.resolve(xhr.response);
                    } else {
                        deferred.reject(xhr);
                    }
                } else {
                    deferred.notify(xhr.readyState);
                }
            };

            xhr.send(data);

            return deferred;
        };

        Context.prototype.post = function (url, data) {
            return this.ajax(url, data, 'POST');
        };

        Context.prototype.getParameterByName = function (url, name) {
            var qmark = url.indexOf('?');

            if (qmark <= 0) {
                return '';
            }

            var regex = new RegExp('[\\?&]' + name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]') + '=([^&#]*)'), results = regex.exec(url.substr(qmark));

            return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
        };

        Context.prototype.getAccessTokenFromRefreshToken = function (resourceId, refreshToken, clientId) {
            var deferred = this.getDeferred(), url = this._authUri + 'oauth2/token', data = 'grant_type=refresh_token&refresh_token=' + encodeURIComponent(refreshToken) + '&client_id=' + encodeURIComponent(clientId) + (resourceId ? '&resource=' + encodeURIComponent(resourceId) : '');

            this.post(url, data).then(function (result) {
                var jsonResult = JSON.parse(result), access_token = {
                    token: jsonResult.access_token,
                    expires_in: new Date((new Date()).getTime() + (jsonResult.expires_in - 300) * 1000)
                };

                this._cacheManager.setAccessToken(clientId, resourceId, access_token);

                // cache most recent refresh token if available.
                deferred.resolve(access_token.token);
            }.bind(this), function (xhr) {
                deferred.reject(new Microsoft.Utility.HttpException(xhr));
            });

            return deferred;
        };

        Context.prototype.isLoginRequired = function (resourceId, clientId) {
            if (!clientId) {
                if (O365Auth.Settings.clientId) {
                    clientId = O365Auth.Settings.clientId;
                } else {
                    throw new Microsoft.Utility.Exception('clientId was not provided nor found in O365Auth.clientId');
                }
            }

            if (resourceId) {
                var access_token = this._cacheManager.getAccessToken(clientId, resourceId);
                if (access_token && access_token.expires_in > new Date()) {
                    return false;
                }
            }

            var refresh_token = this._cacheManager.getRefreshToken(clientId);

            if (refresh_token) {
                return false;
            }

            return true;
        };

        Context.prototype.getAccessToken = function (resourceId, loginHint, clientId, redirectUri) {
            var deferred = this.getDeferred();

            if (!clientId) {
                if (O365Auth.Settings.clientId) {
                    clientId = O365Auth.Settings.clientId;
                } else {
                    deferred.reject(new Microsoft.Utility.Exception('clientId was not provided nor found in O365Auth.clientId'));
                    return deferred;
                }
            }

            if (!redirectUri) {
                redirectUri = this._redirectUri;
            }

            var access_token = this._cacheManager.getAccessToken(clientId, resourceId);

            // five minute time skew.
            if (access_token && access_token.expires_in > new Date()) {
                deferred.resolve(access_token.token);
                return deferred;
            }

            var refresh_token = this._cacheManager.getRefreshToken(clientId);

            if (refresh_token) {
                return this.getAccessTokenFromRefreshToken(resourceId, refresh_token, clientId);
            }

            var authorizationUri = this._authUri + 'oauth2/authorize?response_type=code&resource=' + encodeURIComponent(resourceId) + '&client_id=' + encodeURIComponent(clientId) + '&redirect_uri=' + encodeURIComponent(redirectUri) + (loginHint ? '&login_hint=' + encodeURIComponent(loginHint) : '');

            var onRedirect = function (e) {
                var loadUri = e.url;

                if (loadUri.substr(0, redirectUri.length).toLowerCase() === redirectUri.toLowerCase()) {
                    ref.close();

                    var code = this.getParameterByName(loadUri, 'code'), error = this.getParameterByName(loadUri, 'error_description');

                    if (code) {
                        var url = this._authUri + 'oauth2/token', data = 'grant_type=authorization_code&code=' + code + '&client_id=' + clientId + '&redirect_uri=' + encodeURIComponent(redirectUri);

                        this.post(url, data).then(function (result) {
                            var jsonResult = JSON.parse(result), access_token = {
                                token: jsonResult.access_token,
                                expires_in: new Date((new Date()).getTime() + (jsonResult.expires_in - 300) * 1000)
                            };

                            this._cacheManager.setAccessToken(clientId, resourceId, access_token);
                            this._cacheManager.setIdToken(clientId, jsonResult.id_token);
                            this._cacheManager.setRefreshToken(clientId, jsonResult.refresh_token);

                            deferred.resolve(access_token.token);
                        }.bind(this), function (xhr) {
                            deferred.reject(new Microsoft.Utility.HttpException(xhr));
                        });
                    } else if (error) {
                        deferred.reject(new Microsoft.Utility.Exception(error));
                    }
                }
            }.bind(this);

            var ref = window.open(authorizationUri, '_blank', 'location=yes');

            if (!ref) {
                deferred.reject(new Microsoft.Utility.Exception('The logon dialog was blocked by popup blocker'));
            } else {
                ref.addEventListener('loadstart', onRedirect);

                if (window["tinyHippos"]) {
                    window["__rippleFireEvent"] = onRedirect;
                }
            }

            return deferred;
        };

        Context.prototype.getAccessTokenFn = function (resourceId, loginHint, clientId, redirectUri) {
            return function () {
                return this.getAccessToken(resourceId, loginHint, clientId, redirectUri);
            }.bind(this);
        };

        Context.prototype.getIdToken = function (resourceId, loginHint, clientId, redirectUri) {
            var deferred = this.getDeferred();

            if (!clientId) {
                if (O365Auth.Settings.clientId) {
                    clientId = O365Auth.Settings.clientId;
                } else {
                    deferred.reject(new Microsoft.Utility.Exception('clientId was not provided nor found in O365Auth.clientId'));
                    return deferred;
                }
            }

            if (!redirectUri) {
                redirectUri = this._redirectUri;
            }

            var id_token = this._cacheManager.getIdToken(clientId);

            if (id_token) {
                deferred.resolve(new Token(id_token, this, resourceId, clientId, redirectUri));
            } else {
                this.getAccessToken(resourceId, loginHint, clientId, redirectUri).then(function (value) {
                    var id_token = this._cacheManager.getIdToken(clientId);
                    deferred.resolve(new Token(id_token, this, resourceId, clientId, redirectUri));
                }.bind(this), deferred.reject.bind(deferred));
            }

            return deferred;
        };

        Context.prototype.logOut = function (clientId) {
            var deferred = this.getDeferred(), url = this._authUri + 'oauth2/logout?post_logout_redirect_uri=' + this._redirectUri;

            if (!clientId) {
                if (O365Auth.Settings.clientId) {
                    clientId = O365Auth.Settings.clientId;
                } else {
                    deferred.reject(new Microsoft.Utility.Exception('clientId was not provided nor found in O365Auth.clientId'));
                    return deferred;
                }
            }

            this.ajax(url).then(function (result) {
                deferred.resolve();
            }.bind(this), function (xhr) {
                deferred.reject(new Microsoft.Utility.HttpException(xhr));
            });

            this._cacheManager.clear(clientId);

            return deferred;
        };
        return Context;
    })();
    O365Auth.Context = Context;
})(O365Auth || (O365Auth = {}));
//# sourceMappingURL=o365auth.js.map
﻿var O365Discovery;
(function (O365Discovery) {
    O365Discovery.deferred;

    var Request = (function () {
        function Request(requestUri) {
            this.requestUri = requestUri;
            this.headers = {};
            this.disableCache = false;
        }
        return Request;
    })();
    O365Discovery.Request = Request;

    O365Discovery.capabilityScopes = {
        AllSites: {
            Read: 'AllSites.Read',
            Write: 'AllSites.Write',
            Manage: 'AllSites.Manage',
            FullControl: 'AllSites.FullControl'
        },
        MyFiles: {
            Read: 'MyFiles.Read',
            Write: 'MyFiles.Write'
        },
        user_impersonation: 'user_impersonation',
        full_access: 'full_access',
        Mail: {
            Read: 'Mail.Read',
            Write: 'Mail.Write',
            Sent: 'Mail.Send'
        },
        Calendars: {
            Read: 'Calendars.Read',
            Write: 'Calendars.Write'
        },
        Contacts: {
            Read: 'Contacts.Read',
            Write: 'Contacts.Write'
        }
    };

    (function (AccountType) {
        AccountType[AccountType["MicrosoftAccount"] = 1] = "MicrosoftAccount";
        AccountType[AccountType["OrganizationalId"] = 2] = "OrganizationalId";
    })(O365Discovery.AccountType || (O365Discovery.AccountType = {}));
    var AccountType = O365Discovery.AccountType;

    var ServiceCapability = (function () {
        function ServiceCapability(result) {
            this._result = result;
        }
        Object.defineProperty(ServiceCapability.prototype, "capability", {
            get: function () {
                return this._result.Capability;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ServiceCapability.prototype, "endpointUri", {
            get: function () {
                return this._result.ServiceEndpointUri;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ServiceCapability.prototype, "name", {
            get: function () {
                return this._result.ServiceName;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ServiceCapability.prototype, "resourceId", {
            get: function () {
                return this._result.ServiceResourceId;
            },
            enumerable: true,
            configurable: true
        });
        return ServiceCapability;
    })();
    O365Discovery.ServiceCapability = ServiceCapability;

    var Context = (function () {
        function Context(redirectUri) {
            this._discoveryUri = 'https://api.office.com/discovery/me/';
            this._redirectUri = 'http://localhost/';
            if (!redirectUri) {
                if (O365Auth.Settings.redirectUri) {
                    this._redirectUri = O365Auth.Settings.redirectUri;
                }
            } else {
                this._redirectUri = redirectUri;
            }
        }
        Context.prototype.getDeferred = function () {
            if (O365Discovery.deferred) {
                return O365Discovery.deferred();
            }

            return new Microsoft.Utility.Deferred();
        };

        Context.prototype.ajax = function (request) {
            var deferred = new Microsoft.Utility.Deferred(), xhr = new XMLHttpRequest();

            if (!request.method) {
                request.method = 'GET';
            }

            xhr.open(request.method.toUpperCase(), request.requestUri, true);

            if (request.headers) {
                for (name in request.headers) {
                    var value = request.headers[name];
                    xhr.setRequestHeader(name, request.headers[name]);
                }
            }

            xhr.onreadystatechange = function (e) {
                if (xhr.readyState == 4) {
                    if (xhr.status >= 200 && xhr.status < 300 || xhr.status === 304) {
                        deferred.resolve(xhr.response);
                    } else {
                        deferred.reject(xhr);
                    }
                } else {
                    deferred.notify(xhr.readyState);
                }
            };

            if (request.data) {
                xhr.send(request.data);
            } else {
                xhr.send();
            }

            return deferred;
        };

        Context.prototype.getParameterByName = function (url, name) {
            var qmark = url.indexOf('?');

            if (qmark <= 0) {
                return '';
            }

            var regex = new RegExp('[\\?&]' + name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]') + '=([^&#]*)'), results = regex.exec(url.substr(qmark));

            return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
        };

        Context.prototype.firstSignIn = function (scopes, redirectUri) {
            if (!redirectUri) {
                redirectUri = this._redirectUri;
            }

            var deferred = this.getDeferred(), authorizationUri = this._discoveryUri + 'FirstSignIn?scope=' + scopes + '&redirect_uri=' + encodeURIComponent(redirectUri);

            var onRedirect = function (e) {
                var loadUri = e.url;

                if (loadUri.substr(0, redirectUri.length).toLowerCase() === redirectUri.toLowerCase()) {
                    ref.close();

                    var response = {
                        user_email: this.getParameterByName(loadUri, 'user_email'),
                        account_type: Number(this.getParameterByName(loadUri, 'account_type')),
                        authorization_service: this.getParameterByName(loadUri, 'authorization_service'),
                        token_service: this.getParameterByName(loadUri, 'token_service'),
                        scope: this.getParameterByName(loadUri, 'scope'),
                        unsupported_scope: this.getParameterByName(loadUri, 'unsupported_scope'),
                        discovery_service: this.getParameterByName(loadUri, 'discovery_service'),
                        discovery_resource: this.getParameterByName(loadUri, 'discovery_resource')
                    };

                    deferred.resolve(response);
                }
            }.bind(this);

            var ref = window.open(authorizationUri, '_blank', 'location=yes');

            if (!ref) {
                deferred.reject(new Microsoft.Utility.Exception('The logon dialog was blocked by popup blocker'));
            } else {
                ref.addEventListener('loadstart', onRedirect);

                if (window["tinyHippos"]) {
                    window["__rippleFireEvent"] = onRedirect;
                }
            }

            return deferred;
        };

        Context.prototype.services = function (getAccessTokenFn) {
            var _this = this;
            var deferred = new Microsoft.Utility.Deferred();

            getAccessTokenFn().then((function (value) {
                var request = new Request(_this._discoveryUri + '/services');
                request.headers['Accept'] = 'application/json;odata=verbose';
                request.headers['Authorization'] = 'Bearer ' + value;
                _this.ajax(request).then((function (value) {
                    var parsedData = JSON.parse(value), results = [];

                    parsedData.d.results.forEach(function (v, i, a) {
                        results.push(new ServiceCapability(v));
                    });

                    deferred.resolve(results);
                }).bind(_this), deferred.reject.bind(deferred));
            }).bind(this), deferred.reject.bind(deferred));

            return deferred;
        };

        Context.prototype.allServices = function () {
        };
        return Context;
    })();
    O365Discovery.Context = Context;
})(O365Discovery || (O365Discovery = {}));
//# sourceMappingURL=o365discovery.js.map
﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------
var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var Microsoft;
(function (Microsoft) {
    (function (OutlookServices) {
        (function (Extensions) {
            var ObservableBase = (function () {
                function ObservableBase() {
                    this._changedListeners = [];
                }
                Object.defineProperty(ObservableBase.prototype, "changed", {
                    get: function () {
                        return this._changed;
                    },
                    set: function (value) {
                        var _this = this;
                        this._changed = value;
                        this._changedListeners.forEach((function (value, index, array) {
                            try  {
                                value(_this);
                            } catch (e) {
                            }
                        }).bind(this));
                    },
                    enumerable: true,
                    configurable: true
                });


                ObservableBase.prototype.addChangedListener = function (eventFn) {
                    this._changedListeners.push(eventFn);
                };

                ObservableBase.prototype.removeChangedListener = function (eventFn) {
                    var index = this._changedListeners.indexOf(eventFn);
                    if (index >= 0) {
                        this._changedListeners.splice(index, 1);
                    }
                };
                return ObservableBase;
            })();
            Extensions.ObservableBase = ObservableBase;

            var ObservableCollection = (function (_super) {
                __extends(ObservableCollection, _super);
                function ObservableCollection() {
                    var items = [];
                    for (var _i = 0; _i < (arguments.length - 0); _i++) {
                        items[_i] = arguments[_i + 0];
                    }
                    var _this = this;
                    _super.call(this);
                    this._changedListener = (function (changed) {
                        _this.changed = true;
                    }).bind(this);
                    this._array = items;
                }
                ObservableCollection.prototype.item = function (n) {
                    return this._array[n];
                };

                /**
                * Removes the last element from an array and returns it.
                */
                ObservableCollection.prototype.pop = function () {
                    this.changed = true;
                    var result = this._array.pop();
                    result.removeChangedListener(this._changedListener);
                    return result;
                };

                /**
                * Removes the first element from an array and returns it.
                */
                ObservableCollection.prototype.shift = function () {
                    this.changed = true;
                    var result = this._array.shift();
                    result.removeChangedListener(this._changedListener);
                    return result;
                };

                /**
                * Appends new elements to an array, and returns the new length of the array.
                * @param items New elements of the Array.
                */
                ObservableCollection.prototype.push = function () {
                    var _this = this;
                    var items = [];
                    for (var _i = 0; _i < (arguments.length - 0); _i++) {
                        items[_i] = arguments[_i + 0];
                    }
                    items.forEach((function (value, index, array) {
                        try  {
                            value.addChangedListener(_this._changedListener);
                            _this._array.push(value);
                        } catch (e) {
                        }
                    }).bind(this));
                    this.changed = true;
                    return this._array.length;
                };

                /**
                * Removes elements from an array, returning the deleted elements.
                * @param start The zero-based location in the array from which to start removing elements.
                * @param deleteCount The number of elements to remove.
                * @param items Elements to insert into the array in place of the deleted elements.
                */
                ObservableCollection.prototype.splice = function (start, deleteCount) {
                    var _this = this;
                    var result = this._array.splice(start, deleteCount);
                    result.forEach((function (value, index, array) {
                        try  {
                            value.removeChangedListener(_this._changedListener);
                        } catch (e) {
                        }
                    }).bind(this));
                    this.changed = true;
                    return result;
                };

                /**
                * Inserts new elements at the start of an array.
                * @param items  Elements to insert at the start of the Array.
                */
                ObservableCollection.prototype.unshift = function () {
                    var items = [];
                    for (var _i = 0; _i < (arguments.length - 0); _i++) {
                        items[_i] = arguments[_i + 0];
                    }
                    for (var index = items.length - 1; index >= 0; index--) {
                        try  {
                            items[index].addChangedListener(this._changedListener);
                            this._array.unshift(items[index]);
                        } catch (e) {
                        }
                    }
                    this.changed = true;
                    return this._array.length;
                };

                /**
                * Performs the specified action for each element in an array.
                * @param callbackfn  A function that accepts up to three arguments. forEach calls the callbackfn function one time for each element in the array.
                * @param thisArg  An object to which the this keyword can refer in the callbackfn function. If thisArg is omitted, undefined is used as the this value.
                */
                ObservableCollection.prototype.forEach = function (callbackfn, thisArg) {
                    this._array.forEach(callbackfn, thisArg);
                };

                /**
                * Calls a defined callback function on each element of an array, and returns an array that contains the results.
                * @param callbackfn A function that accepts up to three arguments. The map method calls the callbackfn function one time for each element in the array.
                * @param thisArg An object to which the this keyword can refer in the callbackfn function. If thisArg is omitted, undefined is used as the this value.
                */
                ObservableCollection.prototype.map = function (callbackfn, thisArg) {
                    return this._array.map(callbackfn, thisArg);
                };

                /**
                * Returns the elements of an array that meet the condition specified in a callback function.
                * @param callbackfn A function that accepts up to three arguments. The filter method calls the callbackfn function one time for each element in the array.
                * @param thisArg An object to which the this keyword can refer in the callbackfn function. If thisArg is omitted, undefined is used as the this value.
                */
                ObservableCollection.prototype.filter = function (callbackfn, thisArg) {
                    return this._array.filter(callbackfn, thisArg);
                };

                /**
                * Calls the specified callback function for all the elements in an array. The return value of the callback function is the accumulated result, and is provided as an argument in the next call to the callback function.
                * @param callbackfn A function that accepts up to four arguments. The reduce method calls the callbackfn function one time for each element in the array.
                * @param initialValue If initialValue is specified, it is used as the initial value to start the accumulation. The first call to the callbackfn function provides this value as an argument instead of an array value.
                */
                ObservableCollection.prototype.reduce = function (callbackfn, initialValue) {
                    return this._array.reduce(callbackfn, initialValue);
                };

                /**
                * Calls the specified callback function for all the elements in an array, in descending order. The return value of the callback function is the accumulated result, and is provided as an argument in the next call to the callback function.
                * @param callbackfn A function that accepts up to four arguments. The reduceRight method calls the callbackfn function one time for each element in the array.
                * @param initialValue If initialValue is specified, it is used as the initial value to start the accumulation. The first call to the callbackfn function provides this value as an argument instead of an array value.
                */
                ObservableCollection.prototype.reduceRight = function (callbackfn, initialValue) {
                    return this._array.reduceRight(callbackfn, initialValue);
                };

                Object.defineProperty(ObservableCollection.prototype, "length", {
                    /**
                    * Gets or sets the length of the array. This is a number one higher than the highest element defined in an array.
                    */
                    get: function () {
                        return this._array.length;
                    },
                    enumerable: true,
                    configurable: true
                });
                return ObservableCollection;
            })(ObservableBase);
            Extensions.ObservableCollection = ObservableCollection;

            var Request = (function () {
                function Request(requestUri) {
                    this.requestUri = requestUri;
                    this.headers = {};
                    this.disableCache = false;
                }
                return Request;
            })();
            Extensions.Request = Request;

            var DataContext = (function () {
                function DataContext(serviceRootUri, extraQueryParameters, getAccessTokenFn) {
                    this._noCache = Date.now();
                    this.serviceRootUri = serviceRootUri;
                    this.extraQueryParameters = extraQueryParameters;
                    this._getAccessTokenFn = getAccessTokenFn;
                }
                Object.defineProperty(DataContext.prototype, "serviceRootUri", {
                    get: function () {
                        return this._serviceRootUri;
                    },
                    set: function (value) {
                        if (value.lastIndexOf("/") === value.length - 1) {
                            value = value.substring(0, value.length - 1);
                        }

                        this._serviceRootUri = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                Object.defineProperty(DataContext.prototype, "extraQueryParameters", {
                    get: function () {
                        return this._extraQueryParameters;
                    },
                    set: function (value) {
                        this._extraQueryParameters = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                Object.defineProperty(DataContext.prototype, "disableCache", {
                    get: function () {
                        return this._disableCache;
                    },
                    set: function (value) {
                        this._disableCache = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                Object.defineProperty(DataContext.prototype, "disableCacheOverride", {
                    get: function () {
                        return this._disableCacheOverride;
                    },
                    set: function (value) {
                        this._disableCacheOverride = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                DataContext.prototype.ajax = function (request) {
                    var deferred = new Microsoft.Utility.Deferred();

                    var xhr = new XMLHttpRequest(); xhr.responseType = request.responseType;

                    if (!request.method) {
                        request.method = 'GET';
                    }

                    xhr.open(request.method.toUpperCase(), request.requestUri, true);

                    if (request.headers) {
                        for (name in request.headers) {
                            xhr.setRequestHeader(name, request.headers[name]);
                        }
                    }

                    xhr.onreadystatechange = function (e) {
                        if (xhr.readyState == 4) {
                            if (xhr.status >= 200 && xhr.status < 300 || xhr.status === 304) {
                                deferred.resolve(xhr.response);
                            } else {
                                deferred.reject(xhr);
                            }
                        } else {
                            deferred.notify(xhr.readyState);
                        }
                    };

                    if (request.data) {
                        if ((typeof request.data === 'string') || (Buffer.isBuffer(request.data))) {
                            xhr.send(request.data);
                        } else {
                            xhr.send(JSON.stringify(request.data));
                        }
                    } else {
                        xhr.send();
                    }

                    return deferred;
                };

                DataContext.prototype.read = function (path) {
                    return this.request(new Request(this.serviceRootUri + ((this.serviceRootUri.lastIndexOf('/') != this.serviceRootUri.length - 1) ? '/' : '') + path));
                };

                DataContext.prototype.readUrl = function (url) {
                    return this.request(new Request(url));
                };

                DataContext.prototype.request = function (request) {
                    var _this = this;
                    var deferred;

                    this.augmentRequest(request);

                    if (this._getAccessTokenFn) {
                        deferred = new Microsoft.Utility.Deferred();

                        this._getAccessTokenFn().then((function (token) {
                            request.headers["X-ClientService-ClientTag"] = 'Office 365 API Tools, 1.1.0512';
                            request.headers["Authorization"] = 'Bearer ' + token;
                            _this.ajax(request).then(deferred.resolve.bind(deferred), deferred.reject.bind(deferred));
                        }).bind(this), deferred.reject.bind(deferred));
                    } else {
                        deferred = this.ajax(request);
                    }

                    return deferred;
                };

                DataContext.prototype.augmentRequest = function (request) {
                    if (!request.headers) {
                        request.headers = {};
                    }

                    if (!request.headers['Accept']) {
                        request.headers['Accept'] = 'application/json';
                    }

                    if (!request.headers['Content-Type']) {
                        request.headers['Content-Type'] = 'application/json';
                    }

                    if (this.extraQueryParameters) {
                        request.requestUri += (request.requestUri.indexOf('?') >= 0 ? '&' : '?') + this.extraQueryParameters;
                    }

                    if ((!this._disableCacheOverride && request.disableCache) || (this._disableCacheOverride && this._disableCache)) {
                        request.requestUri += (request.requestUri.indexOf('?') >= 0 ? '&' : '?') + '_=' + this._noCache++;
                    }
                };
                return DataContext;
            })();
            Extensions.DataContext = DataContext;

            var PagedCollection = (function () {
                function PagedCollection(context, path, resultFn, data) {
                    this._context = context;
                    this._path = path;
                    this._resultFn = resultFn;
                    this._data = data;
                }
                Object.defineProperty(PagedCollection.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(PagedCollection.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(PagedCollection.prototype, "currentPage", {
                    get: function () {
                        return this._data;
                    },
                    enumerable: true,
                    configurable: true
                });

                PagedCollection.prototype.getNextPage = function () {
                    var _this = this;
                    var deferred = new Microsoft.Utility.Deferred();

                    if (this.path == null) {
                        deferred.resolve(null);
                        return deferred;
                    }

                    var request = new Request(this.path);

                    request.disableCache = true;

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data), nextLink = (parsedData['odata.nextLink'] === undefined) ? ((parsedData['@odata.nextLink'] === undefined) ? ((parsedData['__next'] === undefined) ? null : parsedData['__next']) : parsedData['@odata.nextLink']) : parsedData['odata.nextLink'];

                        deferred.resolve(new PagedCollection(_this.context, nextLink, _this._resultFn, _this._resultFn(_this.context, parsedData)));
                    }).bind(this), deferred.reject.bind(deferred));

                    return deferred;
                };
                return PagedCollection;
            })();
            Extensions.PagedCollection = PagedCollection;

            var CollectionQuery = (function () {
                function CollectionQuery(context, path, resultFn) {
                    this._context = context;
                    this._path = path;
                    this._resultFn = resultFn;
                }
                Object.defineProperty(CollectionQuery.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(CollectionQuery.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                CollectionQuery.prototype.filter = function (filter) {
                    this.addQuery("$filter=" + filter);
                    return this;
                };

                CollectionQuery.prototype.select = function (selection) {
                    if (typeof selection === 'string') {
                        this.addQuery("$select=" + selection);
                    } else if (Array.isArray(selection)) {
                        this.addQuery("$select=" + selection.join(','));
                    } else {
                        throw new Microsoft.Utility.Exception('\'select\' argument must be string or string[].');
                    }
                    return this;
                };

                CollectionQuery.prototype.expand = function (expand) {
                    if (typeof expand === 'string') {
                        this.addQuery("$expand=" + expand);
                    } else if (Array.isArray(expand)) {
                        this.addQuery("$expand=" + expand.join(','));
                    } else {
                        throw new Microsoft.Utility.Exception('\'expand\' argument must be string or string[].');
                    }
                    return this;
                };

                CollectionQuery.prototype.orderBy = function (orderBy) {
                    if (typeof orderBy === 'string') {
                        this.addQuery("$orderby=" + orderBy);
                    } else if (Array.isArray(orderBy)) {
                        this.addQuery("$orderby=" + orderBy.join(','));
                    } else {
                        throw new Microsoft.Utility.Exception('\'orderBy\' argument must be string or string[].');
                    }
                    return this;
                };

                CollectionQuery.prototype.top = function (top) {
                    this.addQuery("$top=" + top);
                    return this;
                };

                CollectionQuery.prototype.skip = function (skip) {
                    this.addQuery("$skip=" + skip);
                    return this;
                };

                CollectionQuery.prototype.addQuery = function (query) {
                    this._query = (this._query ? this._query + "&" : "") + query;
                    return this;
                };

                Object.defineProperty(CollectionQuery.prototype, "query", {
                    get: function () {
                        return this._query;
                    },
                    set: function (value) {
                        this._query = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                CollectionQuery.prototype.fetch = function () {
                    var path = this.path + (this._query ? (this.path.indexOf('?') < 0 ? '?' : '&') + this._query : "");

                    return new Microsoft.OutlookServices.Extensions.PagedCollection(this.context, path, this._resultFn).getNextPage();
                };

                CollectionQuery.prototype.fetchAll = function (maxItems) {
                    var path = this.path + (this._query ? (this.path.indexOf('?') < 0 ? '?' : '&') + this._query : ""), pagedItems = new Microsoft.OutlookServices.Extensions.PagedCollection(this.context, path, this._resultFn), accumulator = [], deferred = new Microsoft.Utility.Deferred(), recursive = function (nextPagedItems) {
                        if (!nextPagedItems) {
                            deferred.resolve(accumulator);
                        } else {
                            accumulator = accumulator.concat(nextPagedItems.currentPage);

                            if (accumulator.length > maxItems) {
                                accumulator = accumulator.splice(maxItems);
                                deferred.resolve(accumulator);
                            } else {
                                nextPagedItems.getNextPage().then(function (nextPage) {
                                    return recursive(nextPage);
                                }, deferred.reject.bind(deferred));
                            }
                        }
                    };

                    pagedItems.getNextPage().then(function (nextPage) {
                        return recursive(nextPage);
                    }, deferred.reject.bind(deferred));

                    return deferred;
                };
                return CollectionQuery;
            })();
            Extensions.CollectionQuery = CollectionQuery;

            var QueryableSet = (function () {
                function QueryableSet(context, path, entity) {
                    this._context = context;
                    this._path = path;
                    this._entity = entity;
                }
                Object.defineProperty(QueryableSet.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(QueryableSet.prototype, "entity", {
                    get: function () {
                        return this._entity;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(QueryableSet.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                QueryableSet.prototype.getPath = function (prop) {
                    return this._path + '/' + prop;
                };
                return QueryableSet;
            })();
            Extensions.QueryableSet = QueryableSet;

            var RestShallowObjectFetcher = (function () {
                function RestShallowObjectFetcher(context, path) {
                    this._path = path;
                    this._context = context;
                }
                Object.defineProperty(RestShallowObjectFetcher.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(RestShallowObjectFetcher.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                RestShallowObjectFetcher.prototype.getPath = function (prop) {
                    return this._path + '/' + prop;
                };
                return RestShallowObjectFetcher;
            })();
            Extensions.RestShallowObjectFetcher = RestShallowObjectFetcher;

            var ComplexTypeBase = (function (_super) {
                __extends(ComplexTypeBase, _super);
                function ComplexTypeBase() {
                    _super.call(this);
                }
                return ComplexTypeBase;
            })(ObservableBase);
            Extensions.ComplexTypeBase = ComplexTypeBase;

            var EntityBase = (function (_super) {
                __extends(EntityBase, _super);
                function EntityBase(context, path) {
                    _super.call(this);
                    this._path = path;
                    this._context = context;
                }
                Object.defineProperty(EntityBase.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(EntityBase.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                EntityBase.prototype.getPath = function (prop) {
                    return this._path + '/' + prop;
                };
                return EntityBase;
            })(ObservableBase);
            Extensions.EntityBase = EntityBase;

            /*
            std
            */
            function isUndefined(v) {
                return typeof v === 'undefined';
            }
            Extensions.isUndefined = isUndefined;
        })(OutlookServices.Extensions || (OutlookServices.Extensions = {}));
        var Extensions = OutlookServices.Extensions;
    })(Microsoft.OutlookServices || (Microsoft.OutlookServices = {}));
    var OutlookServices = Microsoft.OutlookServices;
})(Microsoft || (Microsoft = {}));

var Microsoft;
(function (Microsoft) {
    (function (OutlookServices) {
        /// <summary>
        /// There are no comments for EntityContainer in the schema.
        /// </summary>
        var Client = (function () {
            function Client(serviceRootUri, getAccessTokenFn) {
                this._context = new Microsoft.OutlookServices.Extensions.DataContext(serviceRootUri, undefined, getAccessTokenFn);
            }
            Object.defineProperty(Client.prototype, "context", {
                get: function () {
                    return this._context;
                },
                enumerable: true,
                configurable: true
            });

            Client.prototype.getPath = function (prop) {
                return this.context.serviceRootUri + '/' + prop;
            };

            Object.defineProperty(Client.prototype, "users", {
                get: function () {
                    if (this._Users === undefined) {
                        this._Users = new Users(this.context, this.getPath('Users'));
                    }
                    return this._Users;
                },
                enumerable: true,
                configurable: true
            });

            /// <summary>
            /// There are no comments for Users in the schema.
            /// </summary>
            Client.prototype.addToUsers = function (user) {
                this.users.addUser(user);
            };

            Object.defineProperty(Client.prototype, "me", {
                get: function () {
                    if (this._Me === undefined) {
                        this._Me = new UserFetcher(this.context, this.getPath("Me"));
                    }
                    return this._Me;
                },
                enumerable: true,
                configurable: true
            });
            return Client;
        })();
        OutlookServices.Client = Client;

        /// <summary>
        /// There are no comments for EmailAddress in the schema.
        /// </summary>
        var EmailAddress = (function (_super) {
            __extends(EmailAddress, _super);
            function EmailAddress(data) {
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.EmailAddress';
                this._NameChanged = false;
                this._AddressChanged = false;

                if (!data) {
                    return;
                }

                this._Name = data.Name;
                this._Address = data.Address;
            }
            Object.defineProperty(EmailAddress.prototype, "name", {
                /// <summary>
                /// There are no comments for Property Name in the schema.
                /// </summary>
                get: function () {
                    return this._Name;
                },
                set: function (value) {
                    if (value !== this._Name) {
                        this._NameChanged = true;
                        this.changed = true;
                    }
                    this._Name = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(EmailAddress.prototype, "nameChanged", {
                get: function () {
                    return this._NameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(EmailAddress.prototype, "address", {
                /// <summary>
                /// There are no comments for Property Address in the schema.
                /// </summary>
                get: function () {
                    return this._Address;
                },
                set: function (value) {
                    if (value !== this._Address) {
                        this._AddressChanged = true;
                        this.changed = true;
                    }
                    this._Address = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(EmailAddress.prototype, "addressChanged", {
                get: function () {
                    return this._AddressChanged;
                },
                enumerable: true,
                configurable: true
            });

            EmailAddress.parseEmailAddress = function (data) {
                if (!data)
                    return null;

                return new EmailAddress(data);
            };

            EmailAddress.parseEmailAddresses = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(EmailAddress.parseEmailAddress(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            EmailAddress.prototype.getRequestBody = function () {
                return {
                    Name: (this.nameChanged && this.name) ? this.name : undefined,
                    Address: (this.addressChanged && this.address) ? this.address : undefined,
                    '@odata.type': this._odataType
                };
            };
            return EmailAddress;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.EmailAddress = EmailAddress;

        /// <summary>
        /// There are no comments for Recipient in the schema.
        /// </summary>
        var Recipient = (function (_super) {
            __extends(Recipient, _super);
            function Recipient(data) {
                var _this = this;
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.Recipient';
                this._EmailAddressChanged = false;
                this._EmailAddressChangedListener = (function (value) {
                    _this._EmailAddressChanged = true;
                    _this.changed = true;
                }).bind(this);

                if (!data) {
                    return;
                }

                this._EmailAddress = EmailAddress.parseEmailAddress(data.EmailAddress);
                if (this._EmailAddress) {
                    this._EmailAddress.addChangedListener(this._EmailAddressChangedListener);
                }
            }
            Object.defineProperty(Recipient.prototype, "emailAddress", {
                /// <summary>
                /// There are no comments for Property EmailAddress in the schema.
                /// </summary>
                get: function () {
                    return this._EmailAddress;
                },
                set: function (value) {
                    if (this._EmailAddress) {
                        this._EmailAddress.removeChangedListener(this._EmailAddressChangedListener);
                    }
                    if (value !== this._EmailAddress) {
                        this._EmailAddressChanged = true;
                        this.changed = true;
                    }
                    if (this._EmailAddress) {
                        this._EmailAddress.addChangedListener(this._EmailAddressChangedListener);
                    }
                    this._EmailAddress = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Recipient.prototype, "emailAddressChanged", {
                get: function () {
                    return this._EmailAddressChanged;
                },
                enumerable: true,
                configurable: true
            });

            Recipient.parseRecipient = function (data) {
                if (!data)
                    return null;

                return new Recipient(data);
            };

            Recipient.parseRecipients = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Recipient.parseRecipient(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            Recipient.prototype.getRequestBody = function () {
                return {
                    EmailAddress: (this.emailAddressChanged && this.emailAddress) ? this.emailAddress.getRequestBody() : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Recipient;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.Recipient = Recipient;

        /// <summary>
        /// There are no comments for Attendee in the schema.
        /// </summary>
        var Attendee = (function (_super) {
            __extends(Attendee, _super);
            function Attendee(data) {
                var _this = this;
                _super.call(this, data);
                this._odataType = 'Microsoft.OutlookServices.Attendee';
                this._StatusChanged = false;
                this._StatusChangedListener = (function (value) {
                    _this._StatusChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._TypeChanged = false;

                if (!data) {
                    return;
                }

                this._Status = ResponseStatus.parseResponseStatus(data.Status);
                if (this._Status) {
                    this._Status.addChangedListener(this._StatusChangedListener);
                }
                this._Type = AttendeeType[data.Type];
            }
            Object.defineProperty(Attendee.prototype, "status", {
                /// <summary>
                /// There are no comments for Property Status in the schema.
                /// </summary>
                get: function () {
                    return this._Status;
                },
                set: function (value) {
                    if (this._Status) {
                        this._Status.removeChangedListener(this._StatusChangedListener);
                    }
                    if (value !== this._Status) {
                        this._StatusChanged = true;
                        this.changed = true;
                    }
                    if (this._Status) {
                        this._Status.addChangedListener(this._StatusChangedListener);
                    }
                    this._Status = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Attendee.prototype, "statusChanged", {
                get: function () {
                    return this._StatusChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Attendee.prototype, "type", {
                /// <summary>
                /// There are no comments for Property Type in the schema.
                /// </summary>
                get: function () {
                    return this._Type;
                },
                set: function (value) {
                    if (value !== this._Type) {
                        this._TypeChanged = true;
                        this.changed = true;
                    }
                    this._Type = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Attendee.prototype, "typeChanged", {
                get: function () {
                    return this._TypeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Attendee.parseAttendee = function (data) {
                if (!data)
                    return null;

                return new Attendee(data);
            };

            Attendee.parseAttendees = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Attendee.parseAttendee(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            Attendee.prototype.getRequestBody = function () {
                return {
                    Status: (this.statusChanged && this.status) ? this.status.getRequestBody() : undefined,
                    Type: (this.typeChanged) ? AttendeeType[this.type] : undefined,
                    EmailAddress: (this.emailAddressChanged && this.emailAddress) ? this.emailAddress.getRequestBody() : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Attendee;
        })(Recipient);
        OutlookServices.Attendee = Attendee;

        /// <summary>
        /// There are no comments for ItemBody in the schema.
        /// </summary>
        var ItemBody = (function (_super) {
            __extends(ItemBody, _super);
            function ItemBody(data) {
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.ItemBody';
                this._ContentTypeChanged = false;
                this._ContentChanged = false;

                if (!data) {
                    return;
                }

                this._ContentType = BodyType[data.ContentType];
                this._Content = data.Content;
            }
            Object.defineProperty(ItemBody.prototype, "contentType", {
                /// <summary>
                /// There are no comments for Property ContentType in the schema.
                /// </summary>
                get: function () {
                    return this._ContentType;
                },
                set: function (value) {
                    if (value !== this._ContentType) {
                        this._ContentTypeChanged = true;
                        this.changed = true;
                    }
                    this._ContentType = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ItemBody.prototype, "contentTypeChanged", {
                get: function () {
                    return this._ContentTypeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(ItemBody.prototype, "content", {
                /// <summary>
                /// There are no comments for Property Content in the schema.
                /// </summary>
                get: function () {
                    return this._Content;
                },
                set: function (value) {
                    if (value !== this._Content) {
                        this._ContentChanged = true;
                        this.changed = true;
                    }
                    this._Content = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ItemBody.prototype, "contentChanged", {
                get: function () {
                    return this._ContentChanged;
                },
                enumerable: true,
                configurable: true
            });

            ItemBody.parseItemBody = function (data) {
                if (!data)
                    return null;

                return new ItemBody(data);
            };

            ItemBody.parseItemBodies = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(ItemBody.parseItemBody(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            ItemBody.prototype.getRequestBody = function () {
                return {
                    ContentType: (this.contentTypeChanged) ? BodyType[this.contentType] : undefined,
                    Content: (this.contentChanged && this.content) ? this.content : undefined,
                    '@odata.type': this._odataType
                };
            };
            return ItemBody;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.ItemBody = ItemBody;

        /// <summary>
        /// There are no comments for Location in the schema.
        /// </summary>
        var Location = (function (_super) {
            __extends(Location, _super);
            function Location(data) {
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.Location';
                this._DisplayNameChanged = false;

                if (!data) {
                    return;
                }

                this._DisplayName = data.DisplayName;
            }
            Object.defineProperty(Location.prototype, "displayName", {
                /// <summary>
                /// There are no comments for Property DisplayName in the schema.
                /// </summary>
                get: function () {
                    return this._DisplayName;
                },
                set: function (value) {
                    if (value !== this._DisplayName) {
                        this._DisplayNameChanged = true;
                        this.changed = true;
                    }
                    this._DisplayName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Location.prototype, "displayNameChanged", {
                get: function () {
                    return this._DisplayNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Location.parseLocation = function (data) {
                if (!data)
                    return null;

                return new Location(data);
            };

            Location.parseLocations = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Location.parseLocation(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            Location.prototype.getRequestBody = function () {
                return {
                    DisplayName: (this.displayNameChanged && this.displayName) ? this.displayName : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Location;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.Location = Location;

        /// <summary>
        /// There are no comments for ResponseStatus in the schema.
        /// </summary>
        var ResponseStatus = (function (_super) {
            __extends(ResponseStatus, _super);
            function ResponseStatus(data) {
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.ResponseStatus';
                this._ResponseChanged = false;
                this._TimeChanged = false;

                if (!data) {
                    return;
                }

                this._Response = ResponseType[data.Response];
                this._Time = (data.Time !== null) ? new Date(data.Time) : null;
            }
            Object.defineProperty(ResponseStatus.prototype, "response", {
                /// <summary>
                /// There are no comments for Property Response in the schema.
                /// </summary>
                get: function () {
                    return this._Response;
                },
                set: function (value) {
                    if (value !== this._Response) {
                        this._ResponseChanged = true;
                        this.changed = true;
                    }
                    this._Response = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ResponseStatus.prototype, "responseChanged", {
                get: function () {
                    return this._ResponseChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(ResponseStatus.prototype, "time", {
                /// <summary>
                /// There are no comments for Property Time in the schema.
                /// </summary>
                get: function () {
                    return this._Time;
                },
                set: function (value) {
                    if (value !== this._Time) {
                        this._TimeChanged = true;
                        this.changed = true;
                    }
                    this._Time = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ResponseStatus.prototype, "timeChanged", {
                get: function () {
                    return this._TimeChanged;
                },
                enumerable: true,
                configurable: true
            });

            ResponseStatus.parseResponseStatus = function (data) {
                if (!data)
                    return null;

                return new ResponseStatus(data);
            };

            ResponseStatus.parseResponseStatuses = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(ResponseStatus.parseResponseStatus(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            ResponseStatus.prototype.getRequestBody = function () {
                return {
                    Response: (this.responseChanged) ? ResponseType[this.response] : undefined,
                    Time: (this.timeChanged && this.time) ? this.time.toString() : undefined,
                    '@odata.type': this._odataType
                };
            };
            return ResponseStatus;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.ResponseStatus = ResponseStatus;

        /// <summary>
        /// There are no comments for PhysicalAddress in the schema.
        /// </summary>
        var PhysicalAddress = (function (_super) {
            __extends(PhysicalAddress, _super);
            function PhysicalAddress(data) {
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.PhysicalAddress';
                this._StreetChanged = false;
                this._CityChanged = false;
                this._StateChanged = false;
                this._CountryOrRegionChanged = false;
                this._PostalCodeChanged = false;

                if (!data) {
                    return;
                }

                this._Street = data.Street;
                this._City = data.City;
                this._State = data.State;
                this._CountryOrRegion = data.CountryOrRegion;
                this._PostalCode = data.PostalCode;
            }
            Object.defineProperty(PhysicalAddress.prototype, "street", {
                /// <summary>
                /// There are no comments for Property Street in the schema.
                /// </summary>
                get: function () {
                    return this._Street;
                },
                set: function (value) {
                    if (value !== this._Street) {
                        this._StreetChanged = true;
                        this.changed = true;
                    }
                    this._Street = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(PhysicalAddress.prototype, "streetChanged", {
                get: function () {
                    return this._StreetChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(PhysicalAddress.prototype, "city", {
                /// <summary>
                /// There are no comments for Property City in the schema.
                /// </summary>
                get: function () {
                    return this._City;
                },
                set: function (value) {
                    if (value !== this._City) {
                        this._CityChanged = true;
                        this.changed = true;
                    }
                    this._City = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(PhysicalAddress.prototype, "cityChanged", {
                get: function () {
                    return this._CityChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(PhysicalAddress.prototype, "state", {
                /// <summary>
                /// There are no comments for Property State in the schema.
                /// </summary>
                get: function () {
                    return this._State;
                },
                set: function (value) {
                    if (value !== this._State) {
                        this._StateChanged = true;
                        this.changed = true;
                    }
                    this._State = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(PhysicalAddress.prototype, "stateChanged", {
                get: function () {
                    return this._StateChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(PhysicalAddress.prototype, "countryOrRegion", {
                /// <summary>
                /// There are no comments for Property CountryOrRegion in the schema.
                /// </summary>
                get: function () {
                    return this._CountryOrRegion;
                },
                set: function (value) {
                    if (value !== this._CountryOrRegion) {
                        this._CountryOrRegionChanged = true;
                        this.changed = true;
                    }
                    this._CountryOrRegion = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(PhysicalAddress.prototype, "countryOrRegionChanged", {
                get: function () {
                    return this._CountryOrRegionChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(PhysicalAddress.prototype, "postalCode", {
                /// <summary>
                /// There are no comments for Property PostalCode in the schema.
                /// </summary>
                get: function () {
                    return this._PostalCode;
                },
                set: function (value) {
                    if (value !== this._PostalCode) {
                        this._PostalCodeChanged = true;
                        this.changed = true;
                    }
                    this._PostalCode = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(PhysicalAddress.prototype, "postalCodeChanged", {
                get: function () {
                    return this._PostalCodeChanged;
                },
                enumerable: true,
                configurable: true
            });

            PhysicalAddress.parsePhysicalAddress = function (data) {
                if (!data)
                    return null;

                return new PhysicalAddress(data);
            };

            PhysicalAddress.parsePhysicalAddresses = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(PhysicalAddress.parsePhysicalAddress(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            PhysicalAddress.prototype.getRequestBody = function () {
                return {
                    Street: (this.streetChanged && this.street) ? this.street : undefined,
                    City: (this.cityChanged && this.city) ? this.city : undefined,
                    State: (this.stateChanged && this.state) ? this.state : undefined,
                    CountryOrRegion: (this.countryOrRegionChanged && this.countryOrRegion) ? this.countryOrRegion : undefined,
                    PostalCode: (this.postalCodeChanged && this.postalCode) ? this.postalCode : undefined,
                    '@odata.type': this._odataType
                };
            };
            return PhysicalAddress;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.PhysicalAddress = PhysicalAddress;

        /// <summary>
        /// There are no comments for RecurrencePattern in the schema.
        /// </summary>
        var RecurrencePattern = (function (_super) {
            __extends(RecurrencePattern, _super);
            function RecurrencePattern(data) {
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.RecurrencePattern';
                this._TypeChanged = false;
                this._IntervalChanged = false;
                this._DayOfMonthChanged = false;
                this._MonthChanged = false;
                this._DaysOfWeek = new Array();
                this._DaysOfWeekChanged = false;
                this._FirstDayOfWeekChanged = false;
                this._IndexChanged = false;

                if (!data) {
                    return;
                }

                this._Type = RecurrencePatternType[data.Type];
                this._Interval = data.Interval;
                this._DayOfMonth = data.DayOfMonth;
                this._Month = data.Month;
                this._DaysOfWeek = data.DaysOfWeek;
                this._FirstDayOfWeek = DayOfWeek[data.FirstDayOfWeek];
                this._Index = WeekIndex[data.Index];
            }
            Object.defineProperty(RecurrencePattern.prototype, "type", {
                /// <summary>
                /// There are no comments for Property Type in the schema.
                /// </summary>
                get: function () {
                    return this._Type;
                },
                set: function (value) {
                    if (value !== this._Type) {
                        this._TypeChanged = true;
                        this.changed = true;
                    }
                    this._Type = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrencePattern.prototype, "typeChanged", {
                get: function () {
                    return this._TypeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrencePattern.prototype, "interval", {
                /// <summary>
                /// There are no comments for Property Interval in the schema.
                /// </summary>
                get: function () {
                    return this._Interval;
                },
                set: function (value) {
                    if (value !== this._Interval) {
                        this._IntervalChanged = true;
                        this.changed = true;
                    }
                    this._Interval = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrencePattern.prototype, "intervalChanged", {
                get: function () {
                    return this._IntervalChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrencePattern.prototype, "dayOfMonth", {
                /// <summary>
                /// There are no comments for Property DayOfMonth in the schema.
                /// </summary>
                get: function () {
                    return this._DayOfMonth;
                },
                set: function (value) {
                    if (value !== this._DayOfMonth) {
                        this._DayOfMonthChanged = true;
                        this.changed = true;
                    }
                    this._DayOfMonth = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrencePattern.prototype, "dayOfMonthChanged", {
                get: function () {
                    return this._DayOfMonthChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrencePattern.prototype, "month", {
                /// <summary>
                /// There are no comments for Property Month in the schema.
                /// </summary>
                get: function () {
                    return this._Month;
                },
                set: function (value) {
                    if (value !== this._Month) {
                        this._MonthChanged = true;
                        this.changed = true;
                    }
                    this._Month = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrencePattern.prototype, "monthChanged", {
                get: function () {
                    return this._MonthChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrencePattern.prototype, "daysOfWeek", {
                /// <summary>
                /// There are no comments for Property DaysOfWeek in the schema.
                /// </summary>
                get: function () {
                    return this._DaysOfWeek;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrencePattern.prototype, "daysOfWeekChanged", {
                get: function () {
                    return this._DaysOfWeekChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrencePattern.prototype, "firstDayOfWeek", {
                /// <summary>
                /// There are no comments for Property FirstDayOfWeek in the schema.
                /// </summary>
                get: function () {
                    return this._FirstDayOfWeek;
                },
                set: function (value) {
                    if (value !== this._FirstDayOfWeek) {
                        this._FirstDayOfWeekChanged = true;
                        this.changed = true;
                    }
                    this._FirstDayOfWeek = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrencePattern.prototype, "firstDayOfWeekChanged", {
                get: function () {
                    return this._FirstDayOfWeekChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrencePattern.prototype, "index", {
                /// <summary>
                /// There are no comments for Property Index in the schema.
                /// </summary>
                get: function () {
                    return this._Index;
                },
                set: function (value) {
                    if (value !== this._Index) {
                        this._IndexChanged = true;
                        this.changed = true;
                    }
                    this._Index = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrencePattern.prototype, "indexChanged", {
                get: function () {
                    return this._IndexChanged;
                },
                enumerable: true,
                configurable: true
            });

            RecurrencePattern.parseRecurrencePattern = function (data) {
                if (!data)
                    return null;

                return new RecurrencePattern(data);
            };

            RecurrencePattern.parseRecurrencePatterns = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(RecurrencePattern.parseRecurrencePattern(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            RecurrencePattern.prototype.getRequestBody = function () {
                return {
                    Type: (this.typeChanged) ? RecurrencePatternType[this.type] : undefined,
                    Interval: (this.intervalChanged && this.interval) ? this.interval : undefined,
                    DayOfMonth: (this.dayOfMonthChanged && this.dayOfMonth) ? this.dayOfMonth : undefined,
                    Month: (this.monthChanged && this.month) ? this.month : undefined,
                    DaysOfWeek: (this.daysOfWeekChanged) ? (function (DaysOfWeek) {
                        if (!DaysOfWeek) {
                            return undefined;
                        }
                        var converted = [];
                        DaysOfWeek.forEach(function (value, index, array) {
                            converted.push(DayOfWeek[value]);
                        });
                        return converted;
                    })(this.daysOfWeek) : undefined,
                    FirstDayOfWeek: (this.firstDayOfWeekChanged) ? DayOfWeek[this.firstDayOfWeek] : undefined,
                    Index: (this.indexChanged) ? WeekIndex[this.index] : undefined,
                    '@odata.type': this._odataType
                };
            };
            return RecurrencePattern;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.RecurrencePattern = RecurrencePattern;

        /// <summary>
        /// There are no comments for RecurrenceRange in the schema.
        /// </summary>
        var RecurrenceRange = (function (_super) {
            __extends(RecurrenceRange, _super);
            function RecurrenceRange(data) {
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.RecurrenceRange';
                this._TypeChanged = false;
                this._StartDateChanged = false;
                this._EndDateChanged = false;
                this._NumberOfOccurrencesChanged = false;

                if (!data) {
                    return;
                }

                this._Type = RecurrenceRangeType[data.Type];
                this._StartDate = (data.StartDate !== null) ? new Date(data.StartDate) : null;
                this._EndDate = (data.EndDate !== null) ? new Date(data.EndDate) : null;
                this._NumberOfOccurrences = data.NumberOfOccurrences;
            }
            Object.defineProperty(RecurrenceRange.prototype, "type", {
                /// <summary>
                /// There are no comments for Property Type in the schema.
                /// </summary>
                get: function () {
                    return this._Type;
                },
                set: function (value) {
                    if (value !== this._Type) {
                        this._TypeChanged = true;
                        this.changed = true;
                    }
                    this._Type = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrenceRange.prototype, "typeChanged", {
                get: function () {
                    return this._TypeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrenceRange.prototype, "startDate", {
                /// <summary>
                /// There are no comments for Property StartDate in the schema.
                /// </summary>
                get: function () {
                    return this._StartDate;
                },
                set: function (value) {
                    if (value !== this._StartDate) {
                        this._StartDateChanged = true;
                        this.changed = true;
                    }
                    this._StartDate = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrenceRange.prototype, "startDateChanged", {
                get: function () {
                    return this._StartDateChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrenceRange.prototype, "endDate", {
                /// <summary>
                /// There are no comments for Property EndDate in the schema.
                /// </summary>
                get: function () {
                    return this._EndDate;
                },
                set: function (value) {
                    if (value !== this._EndDate) {
                        this._EndDateChanged = true;
                        this.changed = true;
                    }
                    this._EndDate = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrenceRange.prototype, "endDateChanged", {
                get: function () {
                    return this._EndDateChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(RecurrenceRange.prototype, "numberOfOccurrences", {
                /// <summary>
                /// There are no comments for Property NumberOfOccurrences in the schema.
                /// </summary>
                get: function () {
                    return this._NumberOfOccurrences;
                },
                set: function (value) {
                    if (value !== this._NumberOfOccurrences) {
                        this._NumberOfOccurrencesChanged = true;
                        this.changed = true;
                    }
                    this._NumberOfOccurrences = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(RecurrenceRange.prototype, "numberOfOccurrencesChanged", {
                get: function () {
                    return this._NumberOfOccurrencesChanged;
                },
                enumerable: true,
                configurable: true
            });

            RecurrenceRange.parseRecurrenceRange = function (data) {
                if (!data)
                    return null;

                return new RecurrenceRange(data);
            };

            RecurrenceRange.parseRecurrenceRanges = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(RecurrenceRange.parseRecurrenceRange(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            RecurrenceRange.prototype.getRequestBody = function () {
                return {
                    Type: (this.typeChanged) ? RecurrenceRangeType[this.type] : undefined,
                    StartDate: (this.startDateChanged && this.startDate) ? this.startDate.toString() : undefined,
                    EndDate: (this.endDateChanged && this.endDate) ? this.endDate.toString() : undefined,
                    NumberOfOccurrences: (this.numberOfOccurrencesChanged && this.numberOfOccurrences) ? this.numberOfOccurrences : undefined,
                    '@odata.type': this._odataType
                };
            };
            return RecurrenceRange;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.RecurrenceRange = RecurrenceRange;

        /// <summary>
        /// There are no comments for PatternedRecurrence in the schema.
        /// </summary>
        var PatternedRecurrence = (function (_super) {
            __extends(PatternedRecurrence, _super);
            function PatternedRecurrence(data) {
                var _this = this;
                _super.call(this);
                this._odataType = 'Microsoft.OutlookServices.PatternedRecurrence';
                this._PatternChanged = false;
                this._PatternChangedListener = (function (value) {
                    _this._PatternChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._RangeChanged = false;
                this._RangeChangedListener = (function (value) {
                    _this._RangeChanged = true;
                    _this.changed = true;
                }).bind(this);

                if (!data) {
                    return;
                }

                this._Pattern = RecurrencePattern.parseRecurrencePattern(data.Pattern);
                if (this._Pattern) {
                    this._Pattern.addChangedListener(this._PatternChangedListener);
                }
                this._Range = RecurrenceRange.parseRecurrenceRange(data.Range);
                if (this._Range) {
                    this._Range.addChangedListener(this._RangeChangedListener);
                }
            }
            Object.defineProperty(PatternedRecurrence.prototype, "pattern", {
                /// <summary>
                /// There are no comments for Property Pattern in the schema.
                /// </summary>
                get: function () {
                    return this._Pattern;
                },
                set: function (value) {
                    if (this._Pattern) {
                        this._Pattern.removeChangedListener(this._PatternChangedListener);
                    }
                    if (value !== this._Pattern) {
                        this._PatternChanged = true;
                        this.changed = true;
                    }
                    if (this._Pattern) {
                        this._Pattern.addChangedListener(this._PatternChangedListener);
                    }
                    this._Pattern = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(PatternedRecurrence.prototype, "patternChanged", {
                get: function () {
                    return this._PatternChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(PatternedRecurrence.prototype, "range", {
                /// <summary>
                /// There are no comments for Property Range in the schema.
                /// </summary>
                get: function () {
                    return this._Range;
                },
                set: function (value) {
                    if (this._Range) {
                        this._Range.removeChangedListener(this._RangeChangedListener);
                    }
                    if (value !== this._Range) {
                        this._RangeChanged = true;
                        this.changed = true;
                    }
                    if (this._Range) {
                        this._Range.addChangedListener(this._RangeChangedListener);
                    }
                    this._Range = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(PatternedRecurrence.prototype, "rangeChanged", {
                get: function () {
                    return this._RangeChanged;
                },
                enumerable: true,
                configurable: true
            });

            PatternedRecurrence.parsePatternedRecurrence = function (data) {
                if (!data)
                    return null;

                return new PatternedRecurrence(data);
            };

            PatternedRecurrence.parsePatternedRecurrences = function (data) {
                var results = new Microsoft.OutlookServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(PatternedRecurrence.parsePatternedRecurrence(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            PatternedRecurrence.prototype.getRequestBody = function () {
                return {
                    Pattern: (this.patternChanged && this.pattern) ? this.pattern.getRequestBody() : undefined,
                    Range: (this.rangeChanged && this.range) ? this.range.getRequestBody() : undefined,
                    '@odata.type': this._odataType
                };
            };
            return PatternedRecurrence;
        })(OutlookServices.Extensions.ComplexTypeBase);
        OutlookServices.PatternedRecurrence = PatternedRecurrence;

        /// <summary>
        /// There are no comments for Entity in the schema.
        /// </summary>
        var EntityFetcher = (function (_super) {
            __extends(EntityFetcher, _super);
            function EntityFetcher(context, path) {
                _super.call(this, context, path);
            }
            return EntityFetcher;
        })(OutlookServices.Extensions.RestShallowObjectFetcher);
        OutlookServices.EntityFetcher = EntityFetcher;

        /// <summary>
        /// There are no comments for Entity in the schema.
        /// </summary>
        var Entity = (function (_super) {
            __extends(Entity, _super);
            function Entity(context, path, data) {
                _super.call(this, context, path);
                this._odataType = 'Microsoft.OutlookServices.Entity';
                this._IdChanged = false;

                if (!data) {
                    return;
                }

                this._Id = data.Id;
            }
            Object.defineProperty(Entity.prototype, "id", {
                /// <summary>
                /// There are no comments for Property Id in the schema.
                /// </summary>
                get: function () {
                    return this._Id;
                },
                set: function (value) {
                    if (value !== this._Id) {
                        this._IdChanged = true;
                        this.changed = true;
                    }
                    this._Id = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Entity.prototype, "idChanged", {
                get: function () {
                    return this._IdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Entity.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Entity.parseEntity(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Entity.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Entity.parseEntity = function (context, path, data) {
                if (!data)
                    return null;

                if (data['@odata.type']) {
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.User')
                        return new User(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Folder')
                        return new Folder(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Item')
                        return new Item(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Message')
                        return new Message(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Attachment')
                        return new Attachment(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.FileAttachment')
                        return new FileAttachment(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.ItemAttachment')
                        return new ItemAttachment(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Calendar')
                        return new Calendar(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.CalendarGroup')
                        return new CalendarGroup(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Event')
                        return new Event(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Contact')
                        return new Contact(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.ContactFolder')
                        return new ContactFolder(context, path, data);
                }

                return new Entity(context, path, data);
            };

            Entity.parseEntities = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Entity.parseEntity(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Entity.prototype.getRequestBody = function () {
                return {
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Entity;
        })(OutlookServices.Extensions.EntityBase);
        OutlookServices.Entity = Entity;

        /// <summary>
        /// There are no comments for User in the schema.
        /// </summary>
        var UserFetcher = (function (_super) {
            __extends(UserFetcher, _super);
            function UserFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(UserFetcher.prototype, "folders", {
                get: function () {
                    if (this._Folders === undefined) {
                        this._Folders = new Microsoft.OutlookServices.Folders(this.context, this.getPath('Folders'));
                    }
                    return this._Folders;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(UserFetcher.prototype, "messages", {
                get: function () {
                    if (this._Messages === undefined) {
                        this._Messages = new Microsoft.OutlookServices.Messages(this.context, this.getPath('Messages'));
                    }
                    return this._Messages;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(UserFetcher.prototype, "rootFolder", {
                /// <summary>
                /// There are no comments for Query Property RootFolder in the schema.
                /// </summary>
                get: function () {
                    if (this._RootFolder === undefined) {
                        this._RootFolder = new FolderFetcher(this.context, this.getPath("RootFolder"));
                    }
                    return this._RootFolder;
                },
                enumerable: true,
                configurable: true
            });

            UserFetcher.prototype.update_rootFolder = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("$links/RootFolder"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Object.defineProperty(UserFetcher.prototype, "calendars", {
                get: function () {
                    if (this._Calendars === undefined) {
                        this._Calendars = new Microsoft.OutlookServices.Calendars(this.context, this.getPath('Calendars'));
                    }
                    return this._Calendars;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(UserFetcher.prototype, "calendar", {
                /// <summary>
                /// There are no comments for Query Property Calendar in the schema.
                /// </summary>
                get: function () {
                    if (this._Calendar === undefined) {
                        this._Calendar = new CalendarFetcher(this.context, this.getPath("Calendar"));
                    }
                    return this._Calendar;
                },
                enumerable: true,
                configurable: true
            });

            UserFetcher.prototype.update_calendar = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("$links/Calendar"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Object.defineProperty(UserFetcher.prototype, "calendarGroups", {
                get: function () {
                    if (this._CalendarGroups === undefined) {
                        this._CalendarGroups = new Microsoft.OutlookServices.CalendarGroups(this.context, this.getPath('CalendarGroups'));
                    }
                    return this._CalendarGroups;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(UserFetcher.prototype, "events", {
                get: function () {
                    if (this._Events === undefined) {
                        this._Events = new Microsoft.OutlookServices.Events(this.context, this.getPath('Events'));
                    }
                    return this._Events;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(UserFetcher.prototype, "calendarView", {
                get: function () {
                    if (this._CalendarView === undefined) {
                        this._CalendarView = new Microsoft.OutlookServices.Events(this.context, this.getPath('CalendarView'));
                    }
                    return this._CalendarView;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(UserFetcher.prototype, "contacts", {
                get: function () {
                    if (this._Contacts === undefined) {
                        this._Contacts = new Microsoft.OutlookServices.Contacts(this.context, this.getPath('Contacts'));
                    }
                    return this._Contacts;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(UserFetcher.prototype, "contactFolders", {
                get: function () {
                    if (this._ContactFolders === undefined) {
                        this._ContactFolders = new Microsoft.OutlookServices.ContactFolders(this.context, this.getPath('ContactFolders'));
                    }
                    return this._ContactFolders;
                },
                enumerable: true,
                configurable: true
            });

            UserFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(User.parseUser(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            UserFetcher.prototype.sendMail = function (Message, SaveToSentItems) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("SendMail"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Message": Message, "SaveToSentItems": SaveToSentItems });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };
            return UserFetcher;
        })(EntityFetcher);
        OutlookServices.UserFetcher = UserFetcher;

        /// <summary>
        /// There are no comments for User in the schema.
        /// </summary>
        var User = (function (_super) {
            __extends(User, _super);
            function User(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.User';
                this._DisplayNameChanged = false;
                this._AliasChanged = false;
                this._MailboxGuidChanged = false;

                if (!data) {
                    return;
                }

                this._DisplayName = data.DisplayName;
                this._Alias = data.Alias;
                this._MailboxGuid = data.MailboxGuid;
            }
            Object.defineProperty(User.prototype, "displayName", {
                /// <summary>
                /// There are no comments for Property DisplayName in the schema.
                /// </summary>
                get: function () {
                    return this._DisplayName;
                },
                set: function (value) {
                    if (value !== this._DisplayName) {
                        this._DisplayNameChanged = true;
                        this.changed = true;
                    }
                    this._DisplayName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(User.prototype, "displayNameChanged", {
                get: function () {
                    return this._DisplayNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "alias", {
                /// <summary>
                /// There are no comments for Property Alias in the schema.
                /// </summary>
                get: function () {
                    return this._Alias;
                },
                set: function (value) {
                    if (value !== this._Alias) {
                        this._AliasChanged = true;
                        this.changed = true;
                    }
                    this._Alias = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(User.prototype, "aliasChanged", {
                get: function () {
                    return this._AliasChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "mailboxGuid", {
                /// <summary>
                /// There are no comments for Property MailboxGuid in the schema.
                /// </summary>
                get: function () {
                    return this._MailboxGuid;
                },
                set: function (value) {
                    if (value !== this._MailboxGuid) {
                        this._MailboxGuidChanged = true;
                        this.changed = true;
                    }
                    this._MailboxGuid = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(User.prototype, "mailboxGuidChanged", {
                get: function () {
                    return this._MailboxGuidChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "folders", {
                get: function () {
                    if (this._Folders === undefined) {
                        this._Folders = new Microsoft.OutlookServices.Folders(this.context, this.getPath('Folders'));
                    }
                    return this._Folders;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "messages", {
                get: function () {
                    if (this._Messages === undefined) {
                        this._Messages = new Microsoft.OutlookServices.Messages(this.context, this.getPath('Messages'));
                    }
                    return this._Messages;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "rootFolder", {
                /// <summary>
                /// There are no comments for Query Property RootFolder in the schema.
                /// </summary>
                get: function () {
                    if (this._RootFolder === undefined) {
                        this._RootFolder = new FolderFetcher(this.context, this.getPath("RootFolder"));
                    }
                    return this._RootFolder;
                },
                enumerable: true,
                configurable: true
            });

            User.prototype.update_rootFolder = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("$links/RootFolder"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Object.defineProperty(User.prototype, "calendars", {
                get: function () {
                    if (this._Calendars === undefined) {
                        this._Calendars = new Microsoft.OutlookServices.Calendars(this.context, this.getPath('Calendars'));
                    }
                    return this._Calendars;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "calendar", {
                /// <summary>
                /// There are no comments for Query Property Calendar in the schema.
                /// </summary>
                get: function () {
                    if (this._Calendar === undefined) {
                        this._Calendar = new CalendarFetcher(this.context, this.getPath("Calendar"));
                    }
                    return this._Calendar;
                },
                enumerable: true,
                configurable: true
            });

            User.prototype.update_calendar = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("$links/Calendar"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Object.defineProperty(User.prototype, "calendarGroups", {
                get: function () {
                    if (this._CalendarGroups === undefined) {
                        this._CalendarGroups = new Microsoft.OutlookServices.CalendarGroups(this.context, this.getPath('CalendarGroups'));
                    }
                    return this._CalendarGroups;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "events", {
                get: function () {
                    if (this._Events === undefined) {
                        this._Events = new Microsoft.OutlookServices.Events(this.context, this.getPath('Events'));
                    }
                    return this._Events;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "calendarView", {
                get: function () {
                    if (this._CalendarView === undefined) {
                        this._CalendarView = new Microsoft.OutlookServices.Events(this.context, this.getPath('CalendarView'));
                    }
                    return this._CalendarView;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "contacts", {
                get: function () {
                    if (this._Contacts === undefined) {
                        this._Contacts = new Microsoft.OutlookServices.Contacts(this.context, this.getPath('Contacts'));
                    }
                    return this._Contacts;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(User.prototype, "contactFolders", {
                get: function () {
                    if (this._ContactFolders === undefined) {
                        this._ContactFolders = new Microsoft.OutlookServices.ContactFolders(this.context, this.getPath('ContactFolders'));
                    }
                    return this._ContactFolders;
                },
                enumerable: true,
                configurable: true
            });

            User.prototype.sendMail = function (Message, SaveToSentItems) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("SendMail"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Message": Message, "SaveToSentItems": SaveToSentItems });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            User.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(User.parseUser(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            User.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            User.parseUser = function (context, path, data) {
                if (!data)
                    return null;

                return new User(context, path, data);
            };

            User.parseUsers = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(User.parseUser(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            User.prototype.getRequestBody = function () {
                return {
                    DisplayName: (this.displayNameChanged && this.displayName) ? this.displayName : undefined,
                    Alias: (this.aliasChanged && this.alias) ? this.alias : undefined,
                    MailboxGuid: (this.mailboxGuidChanged && this.mailboxGuid) ? this.mailboxGuid : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return User;
        })(Entity);
        OutlookServices.User = User;

        /// <summary>
        /// There are no comments for Folder in the schema.
        /// </summary>
        var FolderFetcher = (function (_super) {
            __extends(FolderFetcher, _super);
            function FolderFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(FolderFetcher.prototype, "childFolders", {
                get: function () {
                    if (this._ChildFolders === undefined) {
                        this._ChildFolders = new Microsoft.OutlookServices.Folders(this.context, this.getPath('ChildFolders'));
                    }
                    return this._ChildFolders;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(FolderFetcher.prototype, "messages", {
                get: function () {
                    if (this._Messages === undefined) {
                        this._Messages = new Microsoft.OutlookServices.Messages(this.context, this.getPath('Messages'));
                    }
                    return this._Messages;
                },
                enumerable: true,
                configurable: true
            });

            FolderFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            FolderFetcher.prototype.copy = function (DestinationId) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Copy"));

                request.method = 'POST';
                request.data = JSON.stringify({ "DestinationId": DestinationId });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            FolderFetcher.prototype.move = function (DestinationId) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Move"));

                request.method = 'POST';
                request.data = JSON.stringify({ "DestinationId": DestinationId });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return FolderFetcher;
        })(EntityFetcher);
        OutlookServices.FolderFetcher = FolderFetcher;

        /// <summary>
        /// There are no comments for Folder in the schema.
        /// </summary>
        var Folder = (function (_super) {
            __extends(Folder, _super);
            function Folder(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.Folder';
                this._ParentFolderIdChanged = false;
                this._DisplayNameChanged = false;
                this._ChildFolderCountChanged = false;

                if (!data) {
                    return;
                }

                this._ParentFolderId = data.ParentFolderId;
                this._DisplayName = data.DisplayName;
                this._ChildFolderCount = data.ChildFolderCount;
            }
            Object.defineProperty(Folder.prototype, "parentFolderId", {
                /// <summary>
                /// There are no comments for Property ParentFolderId in the schema.
                /// </summary>
                get: function () {
                    return this._ParentFolderId;
                },
                set: function (value) {
                    if (value !== this._ParentFolderId) {
                        this._ParentFolderIdChanged = true;
                        this.changed = true;
                    }
                    this._ParentFolderId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Folder.prototype, "parentFolderIdChanged", {
                get: function () {
                    return this._ParentFolderIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Folder.prototype, "displayName", {
                /// <summary>
                /// There are no comments for Property DisplayName in the schema.
                /// </summary>
                get: function () {
                    return this._DisplayName;
                },
                set: function (value) {
                    if (value !== this._DisplayName) {
                        this._DisplayNameChanged = true;
                        this.changed = true;
                    }
                    this._DisplayName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Folder.prototype, "displayNameChanged", {
                get: function () {
                    return this._DisplayNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Folder.prototype, "childFolderCount", {
                /// <summary>
                /// There are no comments for Property ChildFolderCount in the schema.
                /// </summary>
                get: function () {
                    return this._ChildFolderCount;
                },
                set: function (value) {
                    if (value !== this._ChildFolderCount) {
                        this._ChildFolderCountChanged = true;
                        this.changed = true;
                    }
                    this._ChildFolderCount = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Folder.prototype, "childFolderCountChanged", {
                get: function () {
                    return this._ChildFolderCountChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Folder.prototype, "childFolders", {
                get: function () {
                    if (this._ChildFolders === undefined) {
                        this._ChildFolders = new Microsoft.OutlookServices.Folders(this.context, this.getPath('ChildFolders'));
                    }
                    return this._ChildFolders;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Folder.prototype, "messages", {
                get: function () {
                    if (this._Messages === undefined) {
                        this._Messages = new Microsoft.OutlookServices.Messages(this.context, this.getPath('Messages'));
                    }
                    return this._Messages;
                },
                enumerable: true,
                configurable: true
            });

            Folder.prototype.copy = function (DestinationId) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Copy"));

                request.method = 'POST';
                request.data = JSON.stringify({ "DestinationId": DestinationId });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            Folder.prototype.move = function (DestinationId) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Move"));

                request.method = 'POST';
                request.data = JSON.stringify({ "DestinationId": DestinationId });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            Folder.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Folder.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Folder.parseFolder = function (context, path, data) {
                if (!data)
                    return null;

                return new Folder(context, path, data);
            };

            Folder.parseFolders = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Folder.parseFolder(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Folder.prototype.getRequestBody = function () {
                return {
                    ParentFolderId: (this.parentFolderIdChanged && this.parentFolderId) ? this.parentFolderId : undefined,
                    DisplayName: (this.displayNameChanged && this.displayName) ? this.displayName : undefined,
                    ChildFolderCount: (this.childFolderCountChanged && this.childFolderCount) ? this.childFolderCount : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Folder;
        })(Entity);
        OutlookServices.Folder = Folder;

        /// <summary>
        /// There are no comments for Item in the schema.
        /// </summary>
        var ItemFetcher = (function (_super) {
            __extends(ItemFetcher, _super);
            function ItemFetcher(context, path) {
                _super.call(this, context, path);
            }
            return ItemFetcher;
        })(EntityFetcher);
        OutlookServices.ItemFetcher = ItemFetcher;

        /// <summary>
        /// There are no comments for Item in the schema.
        /// </summary>
        var Item = (function (_super) {
            __extends(Item, _super);
            function Item(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.Item';
                this._ChangeKeyChanged = false;
                this._Categories = new Array();
                this._CategoriesChanged = false;
                this._DateTimeCreatedChanged = false;
                this._DateTimeLastModifiedChanged = false;

                if (!data) {
                    return;
                }

                this._ChangeKey = data.ChangeKey;
                this._Categories = data.Categories;
                this._DateTimeCreated = (data.DateTimeCreated !== null) ? new Date(data.DateTimeCreated) : null;
                this._DateTimeLastModified = (data.DateTimeLastModified !== null) ? new Date(data.DateTimeLastModified) : null;
            }
            Object.defineProperty(Item.prototype, "changeKey", {
                /// <summary>
                /// There are no comments for Property ChangeKey in the schema.
                /// </summary>
                get: function () {
                    return this._ChangeKey;
                },
                set: function (value) {
                    if (value !== this._ChangeKey) {
                        this._ChangeKeyChanged = true;
                        this.changed = true;
                    }
                    this._ChangeKey = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "changeKeyChanged", {
                get: function () {
                    return this._ChangeKeyChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "categories", {
                /// <summary>
                /// There are no comments for Property Categories in the schema.
                /// </summary>
                get: function () {
                    return this._Categories;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "categoriesChanged", {
                get: function () {
                    return this._CategoriesChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "dateTimeCreated", {
                /// <summary>
                /// There are no comments for Property DateTimeCreated in the schema.
                /// </summary>
                get: function () {
                    return this._DateTimeCreated;
                },
                set: function (value) {
                    if (value !== this._DateTimeCreated) {
                        this._DateTimeCreatedChanged = true;
                        this.changed = true;
                    }
                    this._DateTimeCreated = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "dateTimeCreatedChanged", {
                get: function () {
                    return this._DateTimeCreatedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "dateTimeLastModified", {
                /// <summary>
                /// There are no comments for Property DateTimeLastModified in the schema.
                /// </summary>
                get: function () {
                    return this._DateTimeLastModified;
                },
                set: function (value) {
                    if (value !== this._DateTimeLastModified) {
                        this._DateTimeLastModifiedChanged = true;
                        this.changed = true;
                    }
                    this._DateTimeLastModified = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "dateTimeLastModifiedChanged", {
                get: function () {
                    return this._DateTimeLastModifiedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Item.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Item.parseItem(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Item.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Item.parseItem = function (context, path, data) {
                if (!data)
                    return null;

                if (data['@odata.type']) {
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Message')
                        return new Message(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Event')
                        return new Event(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.Contact')
                        return new Contact(context, path, data);
                }

                return new Item(context, path, data);
            };

            Item.parseItems = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Item.parseItem(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Item.prototype.getRequestBody = function () {
                return {
                    ChangeKey: (this.changeKeyChanged && this.changeKey) ? this.changeKey : undefined,
                    Categories: (this.categoriesChanged && this.categories) ? this.categories : undefined,
                    DateTimeCreated: (this.dateTimeCreatedChanged && this.dateTimeCreated) ? this.dateTimeCreated.toString() : undefined,
                    DateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Item;
        })(Entity);
        OutlookServices.Item = Item;

        /// <summary>
        /// There are no comments for Message in the schema.
        /// </summary>
        var MessageFetcher = (function (_super) {
            __extends(MessageFetcher, _super);
            function MessageFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(MessageFetcher.prototype, "attachments", {
                get: function () {
                    if (this._Attachments === undefined) {
                        this._Attachments = new Microsoft.OutlookServices.Attachments(this.context, this.getPath('Attachments'));
                    }
                    return this._Attachments;
                },
                enumerable: true,
                configurable: true
            });

            MessageFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.copy = function (DestinationId) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Copy"));

                request.method = 'POST';
                request.data = JSON.stringify({ "DestinationId": DestinationId });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.move = function (DestinationId) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Move"));

                request.method = 'POST';
                request.data = JSON.stringify({ "DestinationId": DestinationId });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.createReply = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("CreateReply"));

                request.method = 'POST';

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.createReplyAll = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("CreateReplyAll"));

                request.method = 'POST';

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.createForward = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("CreateForward"));

                request.method = 'POST';

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.reply = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Reply"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.replyAll = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("ReplyAll"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.forward = function (Comment, ToRecipients) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Forward"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment, "ToRecipients": ToRecipients });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            MessageFetcher.prototype.send = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Send"));

                request.method = 'POST';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };
            return MessageFetcher;
        })(ItemFetcher);
        OutlookServices.MessageFetcher = MessageFetcher;

        /// <summary>
        /// There are no comments for Message in the schema.
        /// </summary>
        var Message = (function (_super) {
            __extends(Message, _super);
            function Message(context, path, data) {
                var _this = this;
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.Message';
                this._SubjectChanged = false;
                this._BodyChanged = false;
                this._BodyChangedListener = (function (value) {
                    _this._BodyChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._BodyPreviewChanged = false;
                this._ImportanceChanged = false;
                this._HasAttachmentsChanged = false;
                this._ParentFolderIdChanged = false;
                this._FromChanged = false;
                this._FromChangedListener = (function (value) {
                    _this._FromChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._SenderChanged = false;
                this._SenderChangedListener = (function (value) {
                    _this._SenderChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._ToRecipients = new Microsoft.OutlookServices.Extensions.ObservableCollection();
                this._ToRecipientsChanged = false;
                this._ToRecipientsChangedListener = (function (value) {
                    _this._ToRecipientsChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._CcRecipients = new Microsoft.OutlookServices.Extensions.ObservableCollection();
                this._CcRecipientsChanged = false;
                this._CcRecipientsChangedListener = (function (value) {
                    _this._CcRecipientsChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._BccRecipients = new Microsoft.OutlookServices.Extensions.ObservableCollection();
                this._BccRecipientsChanged = false;
                this._BccRecipientsChangedListener = (function (value) {
                    _this._BccRecipientsChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._ReplyTo = new Microsoft.OutlookServices.Extensions.ObservableCollection();
                this._ReplyToChanged = false;
                this._ReplyToChangedListener = (function (value) {
                    _this._ReplyToChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._ConversationIdChanged = false;
                this._UniqueBodyChanged = false;
                this._UniqueBodyChangedListener = (function (value) {
                    _this._UniqueBodyChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._DateTimeReceivedChanged = false;
                this._DateTimeSentChanged = false;
                this._IsDeliveryReceiptRequestedChanged = false;
                this._IsReadReceiptRequestedChanged = false;
                this._IsDraftChanged = false;
                this._IsReadChanged = false;

                if (!data) {
                    this._ToRecipients.addChangedListener(this._ToRecipientsChangedListener);
                    this._CcRecipients.addChangedListener(this._CcRecipientsChangedListener);
                    this._BccRecipients.addChangedListener(this._BccRecipientsChangedListener);
                    this._ReplyTo.addChangedListener(this._ReplyToChangedListener);
                    return;
                }

                this._Subject = data.Subject;
                this._Body = ItemBody.parseItemBody(data.Body);
                if (this._Body) {
                    this._Body.addChangedListener(this._BodyChangedListener);
                }
                this._BodyPreview = data.BodyPreview;
                this._Importance = Importance[data.Importance];
                this._HasAttachments = data.HasAttachments;
                this._ParentFolderId = data.ParentFolderId;
                this._From = Recipient.parseRecipient(data.From);
                if (this._From) {
                    this._From.addChangedListener(this._FromChangedListener);
                }
                this._Sender = Recipient.parseRecipient(data.Sender);
                if (this._Sender) {
                    this._Sender.addChangedListener(this._SenderChangedListener);
                }
                this._ToRecipients = Recipient.parseRecipients(data.ToRecipients);
                this._ToRecipients.addChangedListener(this._ToRecipientsChangedListener);
                this._CcRecipients = Recipient.parseRecipients(data.CcRecipients);
                this._CcRecipients.addChangedListener(this._CcRecipientsChangedListener);
                this._BccRecipients = Recipient.parseRecipients(data.BccRecipients);
                this._BccRecipients.addChangedListener(this._BccRecipientsChangedListener);
                this._ReplyTo = Recipient.parseRecipients(data.ReplyTo);
                this._ReplyTo.addChangedListener(this._ReplyToChangedListener);
                this._ConversationId = data.ConversationId;
                this._UniqueBody = ItemBody.parseItemBody(data.UniqueBody);
                if (this._UniqueBody) {
                    this._UniqueBody.addChangedListener(this._UniqueBodyChangedListener);
                }
                this._DateTimeReceived = (data.DateTimeReceived !== null) ? new Date(data.DateTimeReceived) : null;
                this._DateTimeSent = (data.DateTimeSent !== null) ? new Date(data.DateTimeSent) : null;
                this._IsDeliveryReceiptRequested = data.IsDeliveryReceiptRequested;
                this._IsReadReceiptRequested = data.IsReadReceiptRequested;
                this._IsDraft = data.IsDraft;
                this._IsRead = data.IsRead;
            }
            Object.defineProperty(Message.prototype, "subject", {
                /// <summary>
                /// There are no comments for Property Subject in the schema.
                /// </summary>
                get: function () {
                    return this._Subject;
                },
                set: function (value) {
                    if (value !== this._Subject) {
                        this._SubjectChanged = true;
                        this.changed = true;
                    }
                    this._Subject = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "subjectChanged", {
                get: function () {
                    return this._SubjectChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "body", {
                /// <summary>
                /// There are no comments for Property Body in the schema.
                /// </summary>
                get: function () {
                    return this._Body;
                },
                set: function (value) {
                    if (this._Body) {
                        this._Body.removeChangedListener(this._BodyChangedListener);
                    }
                    if (value !== this._Body) {
                        this._BodyChanged = true;
                        this.changed = true;
                    }
                    if (this._Body) {
                        this._Body.addChangedListener(this._BodyChangedListener);
                    }
                    this._Body = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "bodyChanged", {
                get: function () {
                    return this._BodyChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "bodyPreview", {
                /// <summary>
                /// There are no comments for Property BodyPreview in the schema.
                /// </summary>
                get: function () {
                    return this._BodyPreview;
                },
                set: function (value) {
                    if (value !== this._BodyPreview) {
                        this._BodyPreviewChanged = true;
                        this.changed = true;
                    }
                    this._BodyPreview = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "bodyPreviewChanged", {
                get: function () {
                    return this._BodyPreviewChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "importance", {
                /// <summary>
                /// There are no comments for Property Importance in the schema.
                /// </summary>
                get: function () {
                    return this._Importance;
                },
                set: function (value) {
                    if (value !== this._Importance) {
                        this._ImportanceChanged = true;
                        this.changed = true;
                    }
                    this._Importance = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "importanceChanged", {
                get: function () {
                    return this._ImportanceChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "hasAttachments", {
                /// <summary>
                /// There are no comments for Property HasAttachments in the schema.
                /// </summary>
                get: function () {
                    return this._HasAttachments;
                },
                set: function (value) {
                    if (value !== this._HasAttachments) {
                        this._HasAttachmentsChanged = true;
                        this.changed = true;
                    }
                    this._HasAttachments = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "hasAttachmentsChanged", {
                get: function () {
                    return this._HasAttachmentsChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "parentFolderId", {
                /// <summary>
                /// There are no comments for Property ParentFolderId in the schema.
                /// </summary>
                get: function () {
                    return this._ParentFolderId;
                },
                set: function (value) {
                    if (value !== this._ParentFolderId) {
                        this._ParentFolderIdChanged = true;
                        this.changed = true;
                    }
                    this._ParentFolderId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "parentFolderIdChanged", {
                get: function () {
                    return this._ParentFolderIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "from", {
                /// <summary>
                /// There are no comments for Property From in the schema.
                /// </summary>
                get: function () {
                    return this._From;
                },
                set: function (value) {
                    if (this._From) {
                        this._From.removeChangedListener(this._FromChangedListener);
                    }
                    if (value !== this._From) {
                        this._FromChanged = true;
                        this.changed = true;
                    }
                    if (this._From) {
                        this._From.addChangedListener(this._FromChangedListener);
                    }
                    this._From = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "fromChanged", {
                get: function () {
                    return this._FromChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "sender", {
                /// <summary>
                /// There are no comments for Property Sender in the schema.
                /// </summary>
                get: function () {
                    return this._Sender;
                },
                set: function (value) {
                    if (this._Sender) {
                        this._Sender.removeChangedListener(this._SenderChangedListener);
                    }
                    if (value !== this._Sender) {
                        this._SenderChanged = true;
                        this.changed = true;
                    }
                    if (this._Sender) {
                        this._Sender.addChangedListener(this._SenderChangedListener);
                    }
                    this._Sender = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "senderChanged", {
                get: function () {
                    return this._SenderChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "toRecipients", {
                /// <summary>
                /// There are no comments for Property ToRecipients in the schema.
                /// </summary>
                get: function () {
                    return this._ToRecipients;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "toRecipientsChanged", {
                get: function () {
                    return this._ToRecipientsChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "ccRecipients", {
                /// <summary>
                /// There are no comments for Property CcRecipients in the schema.
                /// </summary>
                get: function () {
                    return this._CcRecipients;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "ccRecipientsChanged", {
                get: function () {
                    return this._CcRecipientsChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "bccRecipients", {
                /// <summary>
                /// There are no comments for Property BccRecipients in the schema.
                /// </summary>
                get: function () {
                    return this._BccRecipients;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "bccRecipientsChanged", {
                get: function () {
                    return this._BccRecipientsChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "replyTo", {
                /// <summary>
                /// There are no comments for Property ReplyTo in the schema.
                /// </summary>
                get: function () {
                    return this._ReplyTo;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "replyToChanged", {
                get: function () {
                    return this._ReplyToChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "conversationId", {
                /// <summary>
                /// There are no comments for Property ConversationId in the schema.
                /// </summary>
                get: function () {
                    return this._ConversationId;
                },
                set: function (value) {
                    if (value !== this._ConversationId) {
                        this._ConversationIdChanged = true;
                        this.changed = true;
                    }
                    this._ConversationId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "conversationIdChanged", {
                get: function () {
                    return this._ConversationIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "uniqueBody", {
                /// <summary>
                /// There are no comments for Property UniqueBody in the schema.
                /// </summary>
                get: function () {
                    return this._UniqueBody;
                },
                set: function (value) {
                    if (this._UniqueBody) {
                        this._UniqueBody.removeChangedListener(this._UniqueBodyChangedListener);
                    }
                    if (value !== this._UniqueBody) {
                        this._UniqueBodyChanged = true;
                        this.changed = true;
                    }
                    if (this._UniqueBody) {
                        this._UniqueBody.addChangedListener(this._UniqueBodyChangedListener);
                    }
                    this._UniqueBody = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "uniqueBodyChanged", {
                get: function () {
                    return this._UniqueBodyChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "dateTimeReceived", {
                /// <summary>
                /// There are no comments for Property DateTimeReceived in the schema.
                /// </summary>
                get: function () {
                    return this._DateTimeReceived;
                },
                set: function (value) {
                    if (value !== this._DateTimeReceived) {
                        this._DateTimeReceivedChanged = true;
                        this.changed = true;
                    }
                    this._DateTimeReceived = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "dateTimeReceivedChanged", {
                get: function () {
                    return this._DateTimeReceivedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "dateTimeSent", {
                /// <summary>
                /// There are no comments for Property DateTimeSent in the schema.
                /// </summary>
                get: function () {
                    return this._DateTimeSent;
                },
                set: function (value) {
                    if (value !== this._DateTimeSent) {
                        this._DateTimeSentChanged = true;
                        this.changed = true;
                    }
                    this._DateTimeSent = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "dateTimeSentChanged", {
                get: function () {
                    return this._DateTimeSentChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "isDeliveryReceiptRequested", {
                /// <summary>
                /// There are no comments for Property IsDeliveryReceiptRequested in the schema.
                /// </summary>
                get: function () {
                    return this._IsDeliveryReceiptRequested;
                },
                set: function (value) {
                    if (value !== this._IsDeliveryReceiptRequested) {
                        this._IsDeliveryReceiptRequestedChanged = true;
                        this.changed = true;
                    }
                    this._IsDeliveryReceiptRequested = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "isDeliveryReceiptRequestedChanged", {
                get: function () {
                    return this._IsDeliveryReceiptRequestedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "isReadReceiptRequested", {
                /// <summary>
                /// There are no comments for Property IsReadReceiptRequested in the schema.
                /// </summary>
                get: function () {
                    return this._IsReadReceiptRequested;
                },
                set: function (value) {
                    if (value !== this._IsReadReceiptRequested) {
                        this._IsReadReceiptRequestedChanged = true;
                        this.changed = true;
                    }
                    this._IsReadReceiptRequested = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "isReadReceiptRequestedChanged", {
                get: function () {
                    return this._IsReadReceiptRequestedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "isDraft", {
                /// <summary>
                /// There are no comments for Property IsDraft in the schema.
                /// </summary>
                get: function () {
                    return this._IsDraft;
                },
                set: function (value) {
                    if (value !== this._IsDraft) {
                        this._IsDraftChanged = true;
                        this.changed = true;
                    }
                    this._IsDraft = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "isDraftChanged", {
                get: function () {
                    return this._IsDraftChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "isRead", {
                /// <summary>
                /// There are no comments for Property IsRead in the schema.
                /// </summary>
                get: function () {
                    return this._IsRead;
                },
                set: function (value) {
                    if (value !== this._IsRead) {
                        this._IsReadChanged = true;
                        this.changed = true;
                    }
                    this._IsRead = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Message.prototype, "isReadChanged", {
                get: function () {
                    return this._IsReadChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Message.prototype, "attachments", {
                get: function () {
                    if (this._Attachments === undefined) {
                        this._Attachments = new Microsoft.OutlookServices.Attachments(this.context, this.getPath('Attachments'));
                    }
                    return this._Attachments;
                },
                enumerable: true,
                configurable: true
            });

            Message.prototype.copy = function (DestinationId) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Copy"));

                request.method = 'POST';
                request.data = JSON.stringify({ "DestinationId": DestinationId });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.move = function (DestinationId) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Move"));

                request.method = 'POST';
                request.data = JSON.stringify({ "DestinationId": DestinationId });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.createReply = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("CreateReply"));

                request.method = 'POST';

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.createReplyAll = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("CreateReplyAll"));

                request.method = 'POST';

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.createForward = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("CreateForward"));

                request.method = 'POST';

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.reply = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Reply"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.replyAll = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("ReplyAll"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.forward = function (Comment, ToRecipients) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Forward"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment, "ToRecipients": ToRecipients });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.send = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Send"));

                request.method = 'POST';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Message.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Message.parseMessage = function (context, path, data) {
                if (!data)
                    return null;

                return new Message(context, path, data);
            };

            Message.parseMessages = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Message.parseMessage(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Message.prototype.getRequestBody = function () {
                return {
                    Subject: (this.subjectChanged && this.subject) ? this.subject : undefined,
                    Body: (this.bodyChanged && this.body) ? this.body.getRequestBody() : undefined,
                    BodyPreview: (this.bodyPreviewChanged && this.bodyPreview) ? this.bodyPreview : undefined,
                    Importance: (this.importanceChanged) ? Importance[this.importance] : undefined,
                    HasAttachments: (this.hasAttachmentsChanged && this.hasAttachments) ? this.hasAttachments : undefined,
                    ParentFolderId: (this.parentFolderIdChanged && this.parentFolderId) ? this.parentFolderId : undefined,
                    From: (this.fromChanged && this.from) ? this.from.getRequestBody() : undefined,
                    Sender: (this.senderChanged && this.sender) ? this.sender.getRequestBody() : undefined,
                    ToRecipients: (this.toRecipientsChanged) ? (function (ToRecipients) {
                        if (!ToRecipients) {
                            return undefined;
                        }
                        var converted = [];
                        ToRecipients.forEach(function (value, index, array) {
                            converted.push(value.getRequestBody());
                        });
                        return converted;
                    })(this.toRecipients) : undefined,
                    CcRecipients: (this.ccRecipientsChanged) ? (function (CcRecipients) {
                        if (!CcRecipients) {
                            return undefined;
                        }
                        var converted = [];
                        CcRecipients.forEach(function (value, index, array) {
                            converted.push(value.getRequestBody());
                        });
                        return converted;
                    })(this.ccRecipients) : undefined,
                    BccRecipients: (this.bccRecipientsChanged) ? (function (BccRecipients) {
                        if (!BccRecipients) {
                            return undefined;
                        }
                        var converted = [];
                        BccRecipients.forEach(function (value, index, array) {
                            converted.push(value.getRequestBody());
                        });
                        return converted;
                    })(this.bccRecipients) : undefined,
                    ReplyTo: (this.replyToChanged) ? (function (ReplyTo) {
                        if (!ReplyTo) {
                            return undefined;
                        }
                        var converted = [];
                        ReplyTo.forEach(function (value, index, array) {
                            converted.push(value.getRequestBody());
                        });
                        return converted;
                    })(this.replyTo) : undefined,
                    ConversationId: (this.conversationIdChanged && this.conversationId) ? this.conversationId : undefined,
                    UniqueBody: (this.uniqueBodyChanged && this.uniqueBody) ? this.uniqueBody.getRequestBody() : undefined,
                    DateTimeReceived: (this.dateTimeReceivedChanged && this.dateTimeReceived) ? this.dateTimeReceived.toString() : undefined,
                    DateTimeSent: (this.dateTimeSentChanged && this.dateTimeSent) ? this.dateTimeSent.toString() : undefined,
                    IsDeliveryReceiptRequested: (this.isDeliveryReceiptRequestedChanged && this.isDeliveryReceiptRequested) ? this.isDeliveryReceiptRequested : undefined,
                    IsReadReceiptRequested: (this.isReadReceiptRequestedChanged && this.isReadReceiptRequested) ? this.isReadReceiptRequested : undefined,
                    IsDraft: (this.isDraftChanged && this.isDraft) ? this.isDraft : undefined,
                    IsRead: (this.isReadChanged && this.isRead) ? this.isRead : undefined,
                    ChangeKey: (this.changeKeyChanged && this.changeKey) ? this.changeKey : undefined,
                    Categories: (this.categoriesChanged && this.categories) ? this.categories : undefined,
                    DateTimeCreated: (this.dateTimeCreatedChanged && this.dateTimeCreated) ? this.dateTimeCreated.toString() : undefined,
                    DateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Message;
        })(Item);
        OutlookServices.Message = Message;

        /// <summary>
        /// There are no comments for Attachment in the schema.
        /// </summary>
        var AttachmentFetcher = (function (_super) {
            __extends(AttachmentFetcher, _super);
            function AttachmentFetcher(context, path) {
                _super.call(this, context, path);
            }
            return AttachmentFetcher;
        })(EntityFetcher);
        OutlookServices.AttachmentFetcher = AttachmentFetcher;

        /// <summary>
        /// There are no comments for Attachment in the schema.
        /// </summary>
        var Attachment = (function (_super) {
            __extends(Attachment, _super);
            function Attachment(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.Attachment';
                this._NameChanged = false;
                this._ContentTypeChanged = false;
                this._SizeChanged = false;
                this._IsInlineChanged = false;
                this._DateTimeLastModifiedChanged = false;

                if (!data) {
                    return;
                }

                this._Name = data.Name;
                this._ContentType = data.ContentType;
                this._Size = data.Size;
                this._IsInline = data.IsInline;
                this._DateTimeLastModified = (data.DateTimeLastModified !== null) ? new Date(data.DateTimeLastModified) : null;
            }
            Object.defineProperty(Attachment.prototype, "name", {
                /// <summary>
                /// There are no comments for Property Name in the schema.
                /// </summary>
                get: function () {
                    return this._Name;
                },
                set: function (value) {
                    if (value !== this._Name) {
                        this._NameChanged = true;
                        this.changed = true;
                    }
                    this._Name = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Attachment.prototype, "nameChanged", {
                get: function () {
                    return this._NameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Attachment.prototype, "contentType", {
                /// <summary>
                /// There are no comments for Property ContentType in the schema.
                /// </summary>
                get: function () {
                    return this._ContentType;
                },
                set: function (value) {
                    if (value !== this._ContentType) {
                        this._ContentTypeChanged = true;
                        this.changed = true;
                    }
                    this._ContentType = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Attachment.prototype, "contentTypeChanged", {
                get: function () {
                    return this._ContentTypeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Attachment.prototype, "size", {
                /// <summary>
                /// There are no comments for Property Size in the schema.
                /// </summary>
                get: function () {
                    return this._Size;
                },
                set: function (value) {
                    if (value !== this._Size) {
                        this._SizeChanged = true;
                        this.changed = true;
                    }
                    this._Size = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Attachment.prototype, "sizeChanged", {
                get: function () {
                    return this._SizeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Attachment.prototype, "isInline", {
                /// <summary>
                /// There are no comments for Property IsInline in the schema.
                /// </summary>
                get: function () {
                    return this._IsInline;
                },
                set: function (value) {
                    if (value !== this._IsInline) {
                        this._IsInlineChanged = true;
                        this.changed = true;
                    }
                    this._IsInline = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Attachment.prototype, "isInlineChanged", {
                get: function () {
                    return this._IsInlineChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Attachment.prototype, "dateTimeLastModified", {
                /// <summary>
                /// There are no comments for Property DateTimeLastModified in the schema.
                /// </summary>
                get: function () {
                    return this._DateTimeLastModified;
                },
                set: function (value) {
                    if (value !== this._DateTimeLastModified) {
                        this._DateTimeLastModifiedChanged = true;
                        this.changed = true;
                    }
                    this._DateTimeLastModified = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Attachment.prototype, "dateTimeLastModifiedChanged", {
                get: function () {
                    return this._DateTimeLastModifiedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Attachment.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Attachment.parseAttachment(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Attachment.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Attachment.parseAttachment = function (context, path, data) {
                if (!data)
                    return null;

                if (data['@odata.type']) {
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.FileAttachment')
                        return new FileAttachment(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.OutlookServices.ItemAttachment')
                        return new ItemAttachment(context, path, data);
                }

                return new Attachment(context, path, data);
            };

            Attachment.parseAttachments = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Attachment.parseAttachment(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Attachment.prototype.getRequestBody = function () {
                return {
                    Name: (this.nameChanged && this.name) ? this.name : undefined,
                    ContentType: (this.contentTypeChanged && this.contentType) ? this.contentType : undefined,
                    Size: (this.sizeChanged && this.size) ? this.size : undefined,
                    IsInline: (this.isInlineChanged && this.isInline) ? this.isInline : undefined,
                    DateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Attachment;
        })(Entity);
        OutlookServices.Attachment = Attachment;

        /// <summary>
        /// There are no comments for FileAttachment in the schema.
        /// </summary>
        var FileAttachmentFetcher = (function (_super) {
            __extends(FileAttachmentFetcher, _super);
            function FileAttachmentFetcher(context, path) {
                _super.call(this, context, path);
            }
            FileAttachmentFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(FileAttachment.parseFileAttachment(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return FileAttachmentFetcher;
        })(AttachmentFetcher);
        OutlookServices.FileAttachmentFetcher = FileAttachmentFetcher;

        /// <summary>
        /// There are no comments for FileAttachment in the schema.
        /// </summary>
        var FileAttachment = (function (_super) {
            __extends(FileAttachment, _super);
            function FileAttachment(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.FileAttachment';
                this._ContentIdChanged = false;
                this._ContentLocationChanged = false;
                this._IsContactPhotoChanged = false;
                this._ContentBytesChanged = false;

                if (!data) {
                    return;
                }

                this._ContentId = data.ContentId;
                this._ContentLocation = data.ContentLocation;
                this._IsContactPhoto = data.IsContactPhoto;
                this._ContentBytes = data.ContentBytes;
            }
            Object.defineProperty(FileAttachment.prototype, "contentId", {
                /// <summary>
                /// There are no comments for Property ContentId in the schema.
                /// </summary>
                get: function () {
                    return this._ContentId;
                },
                set: function (value) {
                    if (value !== this._ContentId) {
                        this._ContentIdChanged = true;
                        this.changed = true;
                    }
                    this._ContentId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(FileAttachment.prototype, "contentIdChanged", {
                get: function () {
                    return this._ContentIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(FileAttachment.prototype, "contentLocation", {
                /// <summary>
                /// There are no comments for Property ContentLocation in the schema.
                /// </summary>
                get: function () {
                    return this._ContentLocation;
                },
                set: function (value) {
                    if (value !== this._ContentLocation) {
                        this._ContentLocationChanged = true;
                        this.changed = true;
                    }
                    this._ContentLocation = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(FileAttachment.prototype, "contentLocationChanged", {
                get: function () {
                    return this._ContentLocationChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(FileAttachment.prototype, "isContactPhoto", {
                /// <summary>
                /// There are no comments for Property IsContactPhoto in the schema.
                /// </summary>
                get: function () {
                    return this._IsContactPhoto;
                },
                set: function (value) {
                    if (value !== this._IsContactPhoto) {
                        this._IsContactPhotoChanged = true;
                        this.changed = true;
                    }
                    this._IsContactPhoto = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(FileAttachment.prototype, "isContactPhotoChanged", {
                get: function () {
                    return this._IsContactPhotoChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(FileAttachment.prototype, "contentBytes", {
                /// <summary>
                /// There are no comments for Property ContentBytes in the schema.
                /// </summary>
                get: function () {
                    return this._ContentBytes;
                },
                set: function (value) {
                    if (value !== this._ContentBytes) {
                        this._ContentBytesChanged = true;
                        this.changed = true;
                    }
                    this._ContentBytes = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(FileAttachment.prototype, "contentBytesChanged", {
                get: function () {
                    return this._ContentBytesChanged;
                },
                enumerable: true,
                configurable: true
            });

            FileAttachment.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(FileAttachment.parseFileAttachment(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            FileAttachment.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            FileAttachment.parseFileAttachment = function (context, path, data) {
                if (!data)
                    return null;

                return new FileAttachment(context, path, data);
            };

            FileAttachment.parseFileAttachments = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(FileAttachment.parseFileAttachment(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            FileAttachment.prototype.getRequestBody = function () {
                return {
                    ContentId: (this.contentIdChanged && this.contentId) ? this.contentId : undefined,
                    ContentLocation: (this.contentLocationChanged && this.contentLocation) ? this.contentLocation : undefined,
                    IsContactPhoto: (this.isContactPhotoChanged && this.isContactPhoto) ? this.isContactPhoto : undefined,
                    ContentBytes: (this.contentBytesChanged && this.contentBytes) ? this.contentBytes : undefined,
                    Name: (this.nameChanged && this.name) ? this.name : undefined,
                    ContentType: (this.contentTypeChanged && this.contentType) ? this.contentType : undefined,
                    Size: (this.sizeChanged && this.size) ? this.size : undefined,
                    IsInline: (this.isInlineChanged && this.isInline) ? this.isInline : undefined,
                    DateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return FileAttachment;
        })(Attachment);
        OutlookServices.FileAttachment = FileAttachment;

        /// <summary>
        /// There are no comments for ItemAttachment in the schema.
        /// </summary>
        var ItemAttachmentFetcher = (function (_super) {
            __extends(ItemAttachmentFetcher, _super);
            function ItemAttachmentFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(ItemAttachmentFetcher.prototype, "item", {
                /// <summary>
                /// There are no comments for Query Property Item in the schema.
                /// </summary>
                get: function () {
                    if (this._Item === undefined) {
                        this._Item = new ItemFetcher(this.context, this.getPath("Item"));
                    }
                    return this._Item;
                },
                enumerable: true,
                configurable: true
            });

            ItemAttachmentFetcher.prototype.update_item = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("$links/Item"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            ItemAttachmentFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(ItemAttachment.parseItemAttachment(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return ItemAttachmentFetcher;
        })(AttachmentFetcher);
        OutlookServices.ItemAttachmentFetcher = ItemAttachmentFetcher;

        /// <summary>
        /// There are no comments for ItemAttachment in the schema.
        /// </summary>
        var ItemAttachment = (function (_super) {
            __extends(ItemAttachment, _super);
            function ItemAttachment(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.ItemAttachment';

                if (!data) {
                    return;
                }
            }
            Object.defineProperty(ItemAttachment.prototype, "item", {
                /// <summary>
                /// There are no comments for Query Property Item in the schema.
                /// </summary>
                get: function () {
                    if (this._Item === undefined) {
                        this._Item = new ItemFetcher(this.context, this.getPath("Item"));
                    }
                    return this._Item;
                },
                enumerable: true,
                configurable: true
            });

            ItemAttachment.prototype.update_item = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("$links/Item"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            ItemAttachment.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(ItemAttachment.parseItemAttachment(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            ItemAttachment.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            ItemAttachment.parseItemAttachment = function (context, path, data) {
                if (!data)
                    return null;

                return new ItemAttachment(context, path, data);
            };

            ItemAttachment.parseItemAttachments = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(ItemAttachment.parseItemAttachment(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            ItemAttachment.prototype.getRequestBody = function () {
                return {
                    Name: (this.nameChanged && this.name) ? this.name : undefined,
                    ContentType: (this.contentTypeChanged && this.contentType) ? this.contentType : undefined,
                    Size: (this.sizeChanged && this.size) ? this.size : undefined,
                    IsInline: (this.isInlineChanged && this.isInline) ? this.isInline : undefined,
                    DateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return ItemAttachment;
        })(Attachment);
        OutlookServices.ItemAttachment = ItemAttachment;

        /// <summary>
        /// There are no comments for Calendar in the schema.
        /// </summary>
        var CalendarFetcher = (function (_super) {
            __extends(CalendarFetcher, _super);
            function CalendarFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(CalendarFetcher.prototype, "calendarView", {
                get: function () {
                    if (this._CalendarView === undefined) {
                        this._CalendarView = new Microsoft.OutlookServices.Events(this.context, this.getPath('CalendarView'));
                    }
                    return this._CalendarView;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(CalendarFetcher.prototype, "events", {
                get: function () {
                    if (this._Events === undefined) {
                        this._Events = new Microsoft.OutlookServices.Events(this.context, this.getPath('Events'));
                    }
                    return this._Events;
                },
                enumerable: true,
                configurable: true
            });

            CalendarFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Calendar.parseCalendar(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return CalendarFetcher;
        })(EntityFetcher);
        OutlookServices.CalendarFetcher = CalendarFetcher;

        /// <summary>
        /// There are no comments for Calendar in the schema.
        /// </summary>
        var Calendar = (function (_super) {
            __extends(Calendar, _super);
            function Calendar(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.Calendar';
                this._NameChanged = false;
                this._ChangeKeyChanged = false;

                if (!data) {
                    return;
                }

                this._Name = data.Name;
                this._ChangeKey = data.ChangeKey;
            }
            Object.defineProperty(Calendar.prototype, "name", {
                /// <summary>
                /// There are no comments for Property Name in the schema.
                /// </summary>
                get: function () {
                    return this._Name;
                },
                set: function (value) {
                    if (value !== this._Name) {
                        this._NameChanged = true;
                        this.changed = true;
                    }
                    this._Name = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Calendar.prototype, "nameChanged", {
                get: function () {
                    return this._NameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Calendar.prototype, "changeKey", {
                /// <summary>
                /// There are no comments for Property ChangeKey in the schema.
                /// </summary>
                get: function () {
                    return this._ChangeKey;
                },
                set: function (value) {
                    if (value !== this._ChangeKey) {
                        this._ChangeKeyChanged = true;
                        this.changed = true;
                    }
                    this._ChangeKey = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Calendar.prototype, "changeKeyChanged", {
                get: function () {
                    return this._ChangeKeyChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Calendar.prototype, "calendarView", {
                get: function () {
                    if (this._CalendarView === undefined) {
                        this._CalendarView = new Microsoft.OutlookServices.Events(this.context, this.getPath('CalendarView'));
                    }
                    return this._CalendarView;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Calendar.prototype, "events", {
                get: function () {
                    if (this._Events === undefined) {
                        this._Events = new Microsoft.OutlookServices.Events(this.context, this.getPath('Events'));
                    }
                    return this._Events;
                },
                enumerable: true,
                configurable: true
            });

            Calendar.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Calendar.parseCalendar(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Calendar.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Calendar.parseCalendar = function (context, path, data) {
                if (!data)
                    return null;

                return new Calendar(context, path, data);
            };

            Calendar.parseCalendars = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Calendar.parseCalendar(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Calendar.prototype.getRequestBody = function () {
                return {
                    Name: (this.nameChanged && this.name) ? this.name : undefined,
                    ChangeKey: (this.changeKeyChanged && this.changeKey) ? this.changeKey : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Calendar;
        })(Entity);
        OutlookServices.Calendar = Calendar;

        /// <summary>
        /// There are no comments for CalendarGroup in the schema.
        /// </summary>
        var CalendarGroupFetcher = (function (_super) {
            __extends(CalendarGroupFetcher, _super);
            function CalendarGroupFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(CalendarGroupFetcher.prototype, "calendars", {
                get: function () {
                    if (this._Calendars === undefined) {
                        this._Calendars = new Microsoft.OutlookServices.Calendars(this.context, this.getPath('Calendars'));
                    }
                    return this._Calendars;
                },
                enumerable: true,
                configurable: true
            });

            CalendarGroupFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(CalendarGroup.parseCalendarGroup(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return CalendarGroupFetcher;
        })(EntityFetcher);
        OutlookServices.CalendarGroupFetcher = CalendarGroupFetcher;

        /// <summary>
        /// There are no comments for CalendarGroup in the schema.
        /// </summary>
        var CalendarGroup = (function (_super) {
            __extends(CalendarGroup, _super);
            function CalendarGroup(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.CalendarGroup';
                this._NameChanged = false;
                this._ChangeKeyChanged = false;
                this._ClassIdChanged = false;

                if (!data) {
                    return;
                }

                this._Name = data.Name;
                this._ChangeKey = data.ChangeKey;
                this._ClassId = data.ClassId;
            }
            Object.defineProperty(CalendarGroup.prototype, "name", {
                /// <summary>
                /// There are no comments for Property Name in the schema.
                /// </summary>
                get: function () {
                    return this._Name;
                },
                set: function (value) {
                    if (value !== this._Name) {
                        this._NameChanged = true;
                        this.changed = true;
                    }
                    this._Name = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(CalendarGroup.prototype, "nameChanged", {
                get: function () {
                    return this._NameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(CalendarGroup.prototype, "changeKey", {
                /// <summary>
                /// There are no comments for Property ChangeKey in the schema.
                /// </summary>
                get: function () {
                    return this._ChangeKey;
                },
                set: function (value) {
                    if (value !== this._ChangeKey) {
                        this._ChangeKeyChanged = true;
                        this.changed = true;
                    }
                    this._ChangeKey = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(CalendarGroup.prototype, "changeKeyChanged", {
                get: function () {
                    return this._ChangeKeyChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(CalendarGroup.prototype, "classId", {
                /// <summary>
                /// There are no comments for Property ClassId in the schema.
                /// </summary>
                get: function () {
                    return this._ClassId;
                },
                set: function (value) {
                    if (value !== this._ClassId) {
                        this._ClassIdChanged = true;
                        this.changed = true;
                    }
                    this._ClassId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(CalendarGroup.prototype, "classIdChanged", {
                get: function () {
                    return this._ClassIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(CalendarGroup.prototype, "calendars", {
                get: function () {
                    if (this._Calendars === undefined) {
                        this._Calendars = new Microsoft.OutlookServices.Calendars(this.context, this.getPath('Calendars'));
                    }
                    return this._Calendars;
                },
                enumerable: true,
                configurable: true
            });

            CalendarGroup.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(CalendarGroup.parseCalendarGroup(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            CalendarGroup.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            CalendarGroup.parseCalendarGroup = function (context, path, data) {
                if (!data)
                    return null;

                return new CalendarGroup(context, path, data);
            };

            CalendarGroup.parseCalendarGroups = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(CalendarGroup.parseCalendarGroup(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            CalendarGroup.prototype.getRequestBody = function () {
                return {
                    Name: (this.nameChanged && this.name) ? this.name : undefined,
                    ChangeKey: (this.changeKeyChanged && this.changeKey) ? this.changeKey : undefined,
                    ClassId: (this.classIdChanged && this.classId) ? this.classId : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return CalendarGroup;
        })(Entity);
        OutlookServices.CalendarGroup = CalendarGroup;

        /// <summary>
        /// There are no comments for Event in the schema.
        /// </summary>
        var EventFetcher = (function (_super) {
            __extends(EventFetcher, _super);
            function EventFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(EventFetcher.prototype, "attachments", {
                get: function () {
                    if (this._Attachments === undefined) {
                        this._Attachments = new Microsoft.OutlookServices.Attachments(this.context, this.getPath('Attachments'));
                    }
                    return this._Attachments;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(EventFetcher.prototype, "calendar", {
                /// <summary>
                /// There are no comments for Query Property Calendar in the schema.
                /// </summary>
                get: function () {
                    if (this._Calendar === undefined) {
                        this._Calendar = new CalendarFetcher(this.context, this.getPath("Calendar"));
                    }
                    return this._Calendar;
                },
                enumerable: true,
                configurable: true
            });

            EventFetcher.prototype.update_calendar = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("$links/Calendar"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Object.defineProperty(EventFetcher.prototype, "instances", {
                get: function () {
                    if (this._Instances === undefined) {
                        this._Instances = new Microsoft.OutlookServices.Events(this.context, this.getPath('Instances'));
                    }
                    return this._Instances;
                },
                enumerable: true,
                configurable: true
            });

            EventFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Event.parseEvent(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            EventFetcher.prototype.accept = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Accept"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            EventFetcher.prototype.decline = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Decline"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            EventFetcher.prototype.tentativelyAccept = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("TentativelyAccept"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };
            return EventFetcher;
        })(ItemFetcher);
        OutlookServices.EventFetcher = EventFetcher;

        /// <summary>
        /// There are no comments for Event in the schema.
        /// </summary>
        var Event = (function (_super) {
            __extends(Event, _super);
            function Event(context, path, data) {
                var _this = this;
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.Event';
                this._SubjectChanged = false;
                this._BodyChanged = false;
                this._BodyChangedListener = (function (value) {
                    _this._BodyChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._BodyPreviewChanged = false;
                this._ImportanceChanged = false;
                this._HasAttachmentsChanged = false;
                this._StartChanged = false;
                this._EndChanged = false;
                this._LocationChanged = false;
                this._LocationChangedListener = (function (value) {
                    _this._LocationChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._ShowAsChanged = false;
                this._IsAllDayChanged = false;
                this._IsCancelledChanged = false;
                this._IsOrganizerChanged = false;
                this._ResponseRequestedChanged = false;
                this._TypeChanged = false;
                this._SeriesMasterIdChanged = false;
                this._Attendees = new Microsoft.OutlookServices.Extensions.ObservableCollection();
                this._AttendeesChanged = false;
                this._AttendeesChangedListener = (function (value) {
                    _this._AttendeesChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._RecurrenceChanged = false;
                this._RecurrenceChangedListener = (function (value) {
                    _this._RecurrenceChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._OrganizerChanged = false;
                this._OrganizerChangedListener = (function (value) {
                    _this._OrganizerChanged = true;
                    _this.changed = true;
                }).bind(this);

                if (!data) {
                    this._Attendees.addChangedListener(this._AttendeesChangedListener);
                    return;
                }

                this._Subject = data.Subject;
                this._Body = ItemBody.parseItemBody(data.Body);
                if (this._Body) {
                    this._Body.addChangedListener(this._BodyChangedListener);
                }
                this._BodyPreview = data.BodyPreview;
                this._Importance = Importance[data.Importance];
                this._HasAttachments = data.HasAttachments;
                this._Start = (data.Start !== null) ? new Date(data.Start) : null;
                this._End = (data.End !== null) ? new Date(data.End) : null;
                this._Location = Location.parseLocation(data.Location);
                if (this._Location) {
                    this._Location.addChangedListener(this._LocationChangedListener);
                }
                this._ShowAs = FreeBusyStatus[data.ShowAs];
                this._IsAllDay = data.IsAllDay;
                this._IsCancelled = data.IsCancelled;
                this._IsOrganizer = data.IsOrganizer;
                this._ResponseRequested = data.ResponseRequested;
                this._Type = EventType[data.Type];
                this._SeriesMasterId = data.SeriesMasterId;
                this._Attendees = Attendee.parseAttendees(data.Attendees);
                this._Attendees.addChangedListener(this._AttendeesChangedListener);
                this._Recurrence = PatternedRecurrence.parsePatternedRecurrence(data.Recurrence);
                if (this._Recurrence) {
                    this._Recurrence.addChangedListener(this._RecurrenceChangedListener);
                }
                this._Organizer = Recipient.parseRecipient(data.Organizer);
                if (this._Organizer) {
                    this._Organizer.addChangedListener(this._OrganizerChangedListener);
                }
            }
            Object.defineProperty(Event.prototype, "subject", {
                /// <summary>
                /// There are no comments for Property Subject in the schema.
                /// </summary>
                get: function () {
                    return this._Subject;
                },
                set: function (value) {
                    if (value !== this._Subject) {
                        this._SubjectChanged = true;
                        this.changed = true;
                    }
                    this._Subject = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "subjectChanged", {
                get: function () {
                    return this._SubjectChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "body", {
                /// <summary>
                /// There are no comments for Property Body in the schema.
                /// </summary>
                get: function () {
                    return this._Body;
                },
                set: function (value) {
                    if (this._Body) {
                        this._Body.removeChangedListener(this._BodyChangedListener);
                    }
                    if (value !== this._Body) {
                        this._BodyChanged = true;
                        this.changed = true;
                    }
                    if (this._Body) {
                        this._Body.addChangedListener(this._BodyChangedListener);
                    }
                    this._Body = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "bodyChanged", {
                get: function () {
                    return this._BodyChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "bodyPreview", {
                /// <summary>
                /// There are no comments for Property BodyPreview in the schema.
                /// </summary>
                get: function () {
                    return this._BodyPreview;
                },
                set: function (value) {
                    if (value !== this._BodyPreview) {
                        this._BodyPreviewChanged = true;
                        this.changed = true;
                    }
                    this._BodyPreview = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "bodyPreviewChanged", {
                get: function () {
                    return this._BodyPreviewChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "importance", {
                /// <summary>
                /// There are no comments for Property Importance in the schema.
                /// </summary>
                get: function () {
                    return this._Importance;
                },
                set: function (value) {
                    if (value !== this._Importance) {
                        this._ImportanceChanged = true;
                        this.changed = true;
                    }
                    this._Importance = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "importanceChanged", {
                get: function () {
                    return this._ImportanceChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "hasAttachments", {
                /// <summary>
                /// There are no comments for Property HasAttachments in the schema.
                /// </summary>
                get: function () {
                    return this._HasAttachments;
                },
                set: function (value) {
                    if (value !== this._HasAttachments) {
                        this._HasAttachmentsChanged = true;
                        this.changed = true;
                    }
                    this._HasAttachments = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "hasAttachmentsChanged", {
                get: function () {
                    return this._HasAttachmentsChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "start", {
                /// <summary>
                /// There are no comments for Property Start in the schema.
                /// </summary>
                get: function () {
                    return this._Start;
                },
                set: function (value) {
                    if (value !== this._Start) {
                        this._StartChanged = true;
                        this.changed = true;
                    }
                    this._Start = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "startChanged", {
                get: function () {
                    return this._StartChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "end", {
                /// <summary>
                /// There are no comments for Property End in the schema.
                /// </summary>
                get: function () {
                    return this._End;
                },
                set: function (value) {
                    if (value !== this._End) {
                        this._EndChanged = true;
                        this.changed = true;
                    }
                    this._End = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "endChanged", {
                get: function () {
                    return this._EndChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "location", {
                /// <summary>
                /// There are no comments for Property Location in the schema.
                /// </summary>
                get: function () {
                    return this._Location;
                },
                set: function (value) {
                    if (this._Location) {
                        this._Location.removeChangedListener(this._LocationChangedListener);
                    }
                    if (value !== this._Location) {
                        this._LocationChanged = true;
                        this.changed = true;
                    }
                    if (this._Location) {
                        this._Location.addChangedListener(this._LocationChangedListener);
                    }
                    this._Location = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "locationChanged", {
                get: function () {
                    return this._LocationChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "showAs", {
                /// <summary>
                /// There are no comments for Property ShowAs in the schema.
                /// </summary>
                get: function () {
                    return this._ShowAs;
                },
                set: function (value) {
                    if (value !== this._ShowAs) {
                        this._ShowAsChanged = true;
                        this.changed = true;
                    }
                    this._ShowAs = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "showAsChanged", {
                get: function () {
                    return this._ShowAsChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "isAllDay", {
                /// <summary>
                /// There are no comments for Property IsAllDay in the schema.
                /// </summary>
                get: function () {
                    return this._IsAllDay;
                },
                set: function (value) {
                    if (value !== this._IsAllDay) {
                        this._IsAllDayChanged = true;
                        this.changed = true;
                    }
                    this._IsAllDay = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "isAllDayChanged", {
                get: function () {
                    return this._IsAllDayChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "isCancelled", {
                /// <summary>
                /// There are no comments for Property IsCancelled in the schema.
                /// </summary>
                get: function () {
                    return this._IsCancelled;
                },
                set: function (value) {
                    if (value !== this._IsCancelled) {
                        this._IsCancelledChanged = true;
                        this.changed = true;
                    }
                    this._IsCancelled = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "isCancelledChanged", {
                get: function () {
                    return this._IsCancelledChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "isOrganizer", {
                /// <summary>
                /// There are no comments for Property IsOrganizer in the schema.
                /// </summary>
                get: function () {
                    return this._IsOrganizer;
                },
                set: function (value) {
                    if (value !== this._IsOrganizer) {
                        this._IsOrganizerChanged = true;
                        this.changed = true;
                    }
                    this._IsOrganizer = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "isOrganizerChanged", {
                get: function () {
                    return this._IsOrganizerChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "responseRequested", {
                /// <summary>
                /// There are no comments for Property ResponseRequested in the schema.
                /// </summary>
                get: function () {
                    return this._ResponseRequested;
                },
                set: function (value) {
                    if (value !== this._ResponseRequested) {
                        this._ResponseRequestedChanged = true;
                        this.changed = true;
                    }
                    this._ResponseRequested = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "responseRequestedChanged", {
                get: function () {
                    return this._ResponseRequestedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "type", {
                /// <summary>
                /// There are no comments for Property Type in the schema.
                /// </summary>
                get: function () {
                    return this._Type;
                },
                set: function (value) {
                    if (value !== this._Type) {
                        this._TypeChanged = true;
                        this.changed = true;
                    }
                    this._Type = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "typeChanged", {
                get: function () {
                    return this._TypeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "seriesMasterId", {
                /// <summary>
                /// There are no comments for Property SeriesMasterId in the schema.
                /// </summary>
                get: function () {
                    return this._SeriesMasterId;
                },
                set: function (value) {
                    if (value !== this._SeriesMasterId) {
                        this._SeriesMasterIdChanged = true;
                        this.changed = true;
                    }
                    this._SeriesMasterId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "seriesMasterIdChanged", {
                get: function () {
                    return this._SeriesMasterIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "attendees", {
                /// <summary>
                /// There are no comments for Property Attendees in the schema.
                /// </summary>
                get: function () {
                    return this._Attendees;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "attendeesChanged", {
                get: function () {
                    return this._AttendeesChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "recurrence", {
                /// <summary>
                /// There are no comments for Property Recurrence in the schema.
                /// </summary>
                get: function () {
                    return this._Recurrence;
                },
                set: function (value) {
                    if (this._Recurrence) {
                        this._Recurrence.removeChangedListener(this._RecurrenceChangedListener);
                    }
                    if (value !== this._Recurrence) {
                        this._RecurrenceChanged = true;
                        this.changed = true;
                    }
                    if (this._Recurrence) {
                        this._Recurrence.addChangedListener(this._RecurrenceChangedListener);
                    }
                    this._Recurrence = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "recurrenceChanged", {
                get: function () {
                    return this._RecurrenceChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "organizer", {
                /// <summary>
                /// There are no comments for Property Organizer in the schema.
                /// </summary>
                get: function () {
                    return this._Organizer;
                },
                set: function (value) {
                    if (this._Organizer) {
                        this._Organizer.removeChangedListener(this._OrganizerChangedListener);
                    }
                    if (value !== this._Organizer) {
                        this._OrganizerChanged = true;
                        this.changed = true;
                    }
                    if (this._Organizer) {
                        this._Organizer.addChangedListener(this._OrganizerChangedListener);
                    }
                    this._Organizer = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Event.prototype, "organizerChanged", {
                get: function () {
                    return this._OrganizerChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "attachments", {
                get: function () {
                    if (this._Attachments === undefined) {
                        this._Attachments = new Microsoft.OutlookServices.Attachments(this.context, this.getPath('Attachments'));
                    }
                    return this._Attachments;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Event.prototype, "calendar", {
                /// <summary>
                /// There are no comments for Query Property Calendar in the schema.
                /// </summary>
                get: function () {
                    if (this._Calendar === undefined) {
                        this._Calendar = new CalendarFetcher(this.context, this.getPath("Calendar"));
                    }
                    return this._Calendar;
                },
                enumerable: true,
                configurable: true
            });

            Event.prototype.update_calendar = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("$links/Calendar"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Object.defineProperty(Event.prototype, "instances", {
                get: function () {
                    if (this._Instances === undefined) {
                        this._Instances = new Microsoft.OutlookServices.Events(this.context, this.getPath('Instances'));
                    }
                    return this._Instances;
                },
                enumerable: true,
                configurable: true
            });

            Event.prototype.accept = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Accept"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Event.prototype.decline = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("Decline"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Event.prototype.tentativelyAccept = function (Comment) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.getPath("TentativelyAccept"));

                request.method = 'POST';
                request.data = JSON.stringify({ "Comment": Comment });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Event.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Event.parseEvent(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Event.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Event.parseEvent = function (context, path, data) {
                if (!data)
                    return null;

                return new Event(context, path, data);
            };

            Event.parseEvents = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Event.parseEvent(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Event.prototype.getRequestBody = function () {
                return {
                    Subject: (this.subjectChanged && this.subject) ? this.subject : undefined,
                    Body: (this.bodyChanged && this.body) ? this.body.getRequestBody() : undefined,
                    BodyPreview: (this.bodyPreviewChanged && this.bodyPreview) ? this.bodyPreview : undefined,
                    Importance: (this.importanceChanged) ? Importance[this.importance] : undefined,
                    HasAttachments: (this.hasAttachmentsChanged && this.hasAttachments) ? this.hasAttachments : undefined,
                    Start: (this.startChanged && this.start) ? this.start.toString() : undefined,
                    End: (this.endChanged && this.end) ? this.end.toString() : undefined,
                    Location: (this.locationChanged && this.location) ? this.location.getRequestBody() : undefined,
                    ShowAs: (this.showAsChanged) ? FreeBusyStatus[this.showAs] : undefined,
                    IsAllDay: (this.isAllDayChanged && this.isAllDay) ? this.isAllDay : undefined,
                    IsCancelled: (this.isCancelledChanged && this.isCancelled) ? this.isCancelled : undefined,
                    IsOrganizer: (this.isOrganizerChanged && this.isOrganizer) ? this.isOrganizer : undefined,
                    ResponseRequested: (this.responseRequestedChanged && this.responseRequested) ? this.responseRequested : undefined,
                    Type: (this.typeChanged) ? EventType[this.type] : undefined,
                    SeriesMasterId: (this.seriesMasterIdChanged && this.seriesMasterId) ? this.seriesMasterId : undefined,
                    Attendees: (this.attendeesChanged) ? (function (Attendees) {
                        if (!Attendees) {
                            return undefined;
                        }
                        var converted = [];
                        Attendees.forEach(function (value, index, array) {
                            converted.push(value.getRequestBody());
                        });
                        return converted;
                    })(this.attendees) : undefined,
                    Recurrence: (this.recurrenceChanged && this.recurrence) ? this.recurrence.getRequestBody() : undefined,
                    Organizer: (this.organizerChanged && this.organizer) ? this.organizer.getRequestBody() : undefined,
                    ChangeKey: (this.changeKeyChanged && this.changeKey) ? this.changeKey : undefined,
                    Categories: (this.categoriesChanged && this.categories) ? this.categories : undefined,
                    DateTimeCreated: (this.dateTimeCreatedChanged && this.dateTimeCreated) ? this.dateTimeCreated.toString() : undefined,
                    DateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Event;
        })(Item);
        OutlookServices.Event = Event;

        /// <summary>
        /// There are no comments for Contact in the schema.
        /// </summary>
        var ContactFetcher = (function (_super) {
            __extends(ContactFetcher, _super);
            function ContactFetcher(context, path) {
                _super.call(this, context, path);
            }
            ContactFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Contact.parseContact(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return ContactFetcher;
        })(ItemFetcher);
        OutlookServices.ContactFetcher = ContactFetcher;

        /// <summary>
        /// There are no comments for Contact in the schema.
        /// </summary>
        var Contact = (function (_super) {
            __extends(Contact, _super);
            function Contact(context, path, data) {
                var _this = this;
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.Contact';
                this._ParentFolderIdChanged = false;
                this._BirthdayChanged = false;
                this._FileAsChanged = false;
                this._DisplayNameChanged = false;
                this._GivenNameChanged = false;
                this._InitialsChanged = false;
                this._MiddleNameChanged = false;
                this._NickNameChanged = false;
                this._SurnameChanged = false;
                this._TitleChanged = false;
                this._GenerationChanged = false;
                this._EmailAddresses = new Microsoft.OutlookServices.Extensions.ObservableCollection();
                this._EmailAddressesChanged = false;
                this._EmailAddressesChangedListener = (function (value) {
                    _this._EmailAddressesChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._ImAddresses = new Array();
                this._ImAddressesChanged = false;
                this._JobTitleChanged = false;
                this._CompanyNameChanged = false;
                this._DepartmentChanged = false;
                this._OfficeLocationChanged = false;
                this._ProfessionChanged = false;
                this._BusinessHomePageChanged = false;
                this._AssistantNameChanged = false;
                this._ManagerChanged = false;
                this._HomePhones = new Array();
                this._HomePhonesChanged = false;
                this._BusinessPhones = new Array();
                this._BusinessPhonesChanged = false;
                this._MobilePhone1Changed = false;
                this._HomeAddressChanged = false;
                this._HomeAddressChangedListener = (function (value) {
                    _this._HomeAddressChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._BusinessAddressChanged = false;
                this._BusinessAddressChangedListener = (function (value) {
                    _this._BusinessAddressChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._OtherAddressChanged = false;
                this._OtherAddressChangedListener = (function (value) {
                    _this._OtherAddressChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._YomiCompanyNameChanged = false;
                this._YomiGivenNameChanged = false;
                this._YomiSurnameChanged = false;

                if (!data) {
                    this._EmailAddresses.addChangedListener(this._EmailAddressesChangedListener);
                    return;
                }

                this._ParentFolderId = data.ParentFolderId;
                this._Birthday = (data.Birthday !== null) ? new Date(data.Birthday) : null;
                this._FileAs = data.FileAs;
                this._DisplayName = data.DisplayName;
                this._GivenName = data.GivenName;
                this._Initials = data.Initials;
                this._MiddleName = data.MiddleName;
                this._NickName = data.NickName;
                this._Surname = data.Surname;
                this._Title = data.Title;
                this._Generation = data.Generation;
                this._EmailAddresses = EmailAddress.parseEmailAddresses(data.EmailAddresses);
                this._EmailAddresses.addChangedListener(this._EmailAddressesChangedListener);
                this._ImAddresses = data.ImAddresses;
                this._JobTitle = data.JobTitle;
                this._CompanyName = data.CompanyName;
                this._Department = data.Department;
                this._OfficeLocation = data.OfficeLocation;
                this._Profession = data.Profession;
                this._BusinessHomePage = data.BusinessHomePage;
                this._AssistantName = data.AssistantName;
                this._Manager = data.Manager;
                this._HomePhones = data.HomePhones;
                this._BusinessPhones = data.BusinessPhones;
                this._MobilePhone1 = data.MobilePhone1;
                this._HomeAddress = PhysicalAddress.parsePhysicalAddress(data.HomeAddress);
                if (this._HomeAddress) {
                    this._HomeAddress.addChangedListener(this._HomeAddressChangedListener);
                }
                this._BusinessAddress = PhysicalAddress.parsePhysicalAddress(data.BusinessAddress);
                if (this._BusinessAddress) {
                    this._BusinessAddress.addChangedListener(this._BusinessAddressChangedListener);
                }
                this._OtherAddress = PhysicalAddress.parsePhysicalAddress(data.OtherAddress);
                if (this._OtherAddress) {
                    this._OtherAddress.addChangedListener(this._OtherAddressChangedListener);
                }
                this._YomiCompanyName = data.YomiCompanyName;
                this._YomiGivenName = data.YomiGivenName;
                this._YomiSurname = data.YomiSurname;
            }
            Object.defineProperty(Contact.prototype, "parentFolderId", {
                /// <summary>
                /// There are no comments for Property ParentFolderId in the schema.
                /// </summary>
                get: function () {
                    return this._ParentFolderId;
                },
                set: function (value) {
                    if (value !== this._ParentFolderId) {
                        this._ParentFolderIdChanged = true;
                        this.changed = true;
                    }
                    this._ParentFolderId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "parentFolderIdChanged", {
                get: function () {
                    return this._ParentFolderIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "birthday", {
                /// <summary>
                /// There are no comments for Property Birthday in the schema.
                /// </summary>
                get: function () {
                    return this._Birthday;
                },
                set: function (value) {
                    if (value !== this._Birthday) {
                        this._BirthdayChanged = true;
                        this.changed = true;
                    }
                    this._Birthday = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "birthdayChanged", {
                get: function () {
                    return this._BirthdayChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "fileAs", {
                /// <summary>
                /// There are no comments for Property FileAs in the schema.
                /// </summary>
                get: function () {
                    return this._FileAs;
                },
                set: function (value) {
                    if (value !== this._FileAs) {
                        this._FileAsChanged = true;
                        this.changed = true;
                    }
                    this._FileAs = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "fileAsChanged", {
                get: function () {
                    return this._FileAsChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "displayName", {
                /// <summary>
                /// There are no comments for Property DisplayName in the schema.
                /// </summary>
                get: function () {
                    return this._DisplayName;
                },
                set: function (value) {
                    if (value !== this._DisplayName) {
                        this._DisplayNameChanged = true;
                        this.changed = true;
                    }
                    this._DisplayName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "displayNameChanged", {
                get: function () {
                    return this._DisplayNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "givenName", {
                /// <summary>
                /// There are no comments for Property GivenName in the schema.
                /// </summary>
                get: function () {
                    return this._GivenName;
                },
                set: function (value) {
                    if (value !== this._GivenName) {
                        this._GivenNameChanged = true;
                        this.changed = true;
                    }
                    this._GivenName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "givenNameChanged", {
                get: function () {
                    return this._GivenNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "initials", {
                /// <summary>
                /// There are no comments for Property Initials in the schema.
                /// </summary>
                get: function () {
                    return this._Initials;
                },
                set: function (value) {
                    if (value !== this._Initials) {
                        this._InitialsChanged = true;
                        this.changed = true;
                    }
                    this._Initials = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "initialsChanged", {
                get: function () {
                    return this._InitialsChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "middleName", {
                /// <summary>
                /// There are no comments for Property MiddleName in the schema.
                /// </summary>
                get: function () {
                    return this._MiddleName;
                },
                set: function (value) {
                    if (value !== this._MiddleName) {
                        this._MiddleNameChanged = true;
                        this.changed = true;
                    }
                    this._MiddleName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "middleNameChanged", {
                get: function () {
                    return this._MiddleNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "nickName", {
                /// <summary>
                /// There are no comments for Property NickName in the schema.
                /// </summary>
                get: function () {
                    return this._NickName;
                },
                set: function (value) {
                    if (value !== this._NickName) {
                        this._NickNameChanged = true;
                        this.changed = true;
                    }
                    this._NickName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "nickNameChanged", {
                get: function () {
                    return this._NickNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "surname", {
                /// <summary>
                /// There are no comments for Property Surname in the schema.
                /// </summary>
                get: function () {
                    return this._Surname;
                },
                set: function (value) {
                    if (value !== this._Surname) {
                        this._SurnameChanged = true;
                        this.changed = true;
                    }
                    this._Surname = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "surnameChanged", {
                get: function () {
                    return this._SurnameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "title", {
                /// <summary>
                /// There are no comments for Property Title in the schema.
                /// </summary>
                get: function () {
                    return this._Title;
                },
                set: function (value) {
                    if (value !== this._Title) {
                        this._TitleChanged = true;
                        this.changed = true;
                    }
                    this._Title = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "titleChanged", {
                get: function () {
                    return this._TitleChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "generation", {
                /// <summary>
                /// There are no comments for Property Generation in the schema.
                /// </summary>
                get: function () {
                    return this._Generation;
                },
                set: function (value) {
                    if (value !== this._Generation) {
                        this._GenerationChanged = true;
                        this.changed = true;
                    }
                    this._Generation = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "generationChanged", {
                get: function () {
                    return this._GenerationChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "emailAddresses", {
                /// <summary>
                /// There are no comments for Property EmailAddresses in the schema.
                /// </summary>
                get: function () {
                    return this._EmailAddresses;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "emailAddressesChanged", {
                get: function () {
                    return this._EmailAddressesChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "imAddresses", {
                /// <summary>
                /// There are no comments for Property ImAddresses in the schema.
                /// </summary>
                get: function () {
                    return this._ImAddresses;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "imAddressesChanged", {
                get: function () {
                    return this._ImAddressesChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "jobTitle", {
                /// <summary>
                /// There are no comments for Property JobTitle in the schema.
                /// </summary>
                get: function () {
                    return this._JobTitle;
                },
                set: function (value) {
                    if (value !== this._JobTitle) {
                        this._JobTitleChanged = true;
                        this.changed = true;
                    }
                    this._JobTitle = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "jobTitleChanged", {
                get: function () {
                    return this._JobTitleChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "companyName", {
                /// <summary>
                /// There are no comments for Property CompanyName in the schema.
                /// </summary>
                get: function () {
                    return this._CompanyName;
                },
                set: function (value) {
                    if (value !== this._CompanyName) {
                        this._CompanyNameChanged = true;
                        this.changed = true;
                    }
                    this._CompanyName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "companyNameChanged", {
                get: function () {
                    return this._CompanyNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "department", {
                /// <summary>
                /// There are no comments for Property Department in the schema.
                /// </summary>
                get: function () {
                    return this._Department;
                },
                set: function (value) {
                    if (value !== this._Department) {
                        this._DepartmentChanged = true;
                        this.changed = true;
                    }
                    this._Department = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "departmentChanged", {
                get: function () {
                    return this._DepartmentChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "officeLocation", {
                /// <summary>
                /// There are no comments for Property OfficeLocation in the schema.
                /// </summary>
                get: function () {
                    return this._OfficeLocation;
                },
                set: function (value) {
                    if (value !== this._OfficeLocation) {
                        this._OfficeLocationChanged = true;
                        this.changed = true;
                    }
                    this._OfficeLocation = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "officeLocationChanged", {
                get: function () {
                    return this._OfficeLocationChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "profession", {
                /// <summary>
                /// There are no comments for Property Profession in the schema.
                /// </summary>
                get: function () {
                    return this._Profession;
                },
                set: function (value) {
                    if (value !== this._Profession) {
                        this._ProfessionChanged = true;
                        this.changed = true;
                    }
                    this._Profession = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "professionChanged", {
                get: function () {
                    return this._ProfessionChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "businessHomePage", {
                /// <summary>
                /// There are no comments for Property BusinessHomePage in the schema.
                /// </summary>
                get: function () {
                    return this._BusinessHomePage;
                },
                set: function (value) {
                    if (value !== this._BusinessHomePage) {
                        this._BusinessHomePageChanged = true;
                        this.changed = true;
                    }
                    this._BusinessHomePage = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "businessHomePageChanged", {
                get: function () {
                    return this._BusinessHomePageChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "assistantName", {
                /// <summary>
                /// There are no comments for Property AssistantName in the schema.
                /// </summary>
                get: function () {
                    return this._AssistantName;
                },
                set: function (value) {
                    if (value !== this._AssistantName) {
                        this._AssistantNameChanged = true;
                        this.changed = true;
                    }
                    this._AssistantName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "assistantNameChanged", {
                get: function () {
                    return this._AssistantNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "manager", {
                /// <summary>
                /// There are no comments for Property Manager in the schema.
                /// </summary>
                get: function () {
                    return this._Manager;
                },
                set: function (value) {
                    if (value !== this._Manager) {
                        this._ManagerChanged = true;
                        this.changed = true;
                    }
                    this._Manager = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "managerChanged", {
                get: function () {
                    return this._ManagerChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "homePhones", {
                /// <summary>
                /// There are no comments for Property HomePhones in the schema.
                /// </summary>
                get: function () {
                    return this._HomePhones;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "homePhonesChanged", {
                get: function () {
                    return this._HomePhonesChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "businessPhones", {
                /// <summary>
                /// There are no comments for Property BusinessPhones in the schema.
                /// </summary>
                get: function () {
                    return this._BusinessPhones;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "businessPhonesChanged", {
                get: function () {
                    return this._BusinessPhonesChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "mobilePhone1", {
                /// <summary>
                /// There are no comments for Property MobilePhone1 in the schema.
                /// </summary>
                get: function () {
                    return this._MobilePhone1;
                },
                set: function (value) {
                    if (value !== this._MobilePhone1) {
                        this._MobilePhone1Changed = true;
                        this.changed = true;
                    }
                    this._MobilePhone1 = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "mobilePhone1Changed", {
                get: function () {
                    return this._MobilePhone1Changed;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "homeAddress", {
                /// <summary>
                /// There are no comments for Property HomeAddress in the schema.
                /// </summary>
                get: function () {
                    return this._HomeAddress;
                },
                set: function (value) {
                    if (this._HomeAddress) {
                        this._HomeAddress.removeChangedListener(this._HomeAddressChangedListener);
                    }
                    if (value !== this._HomeAddress) {
                        this._HomeAddressChanged = true;
                        this.changed = true;
                    }
                    if (this._HomeAddress) {
                        this._HomeAddress.addChangedListener(this._HomeAddressChangedListener);
                    }
                    this._HomeAddress = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "homeAddressChanged", {
                get: function () {
                    return this._HomeAddressChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "businessAddress", {
                /// <summary>
                /// There are no comments for Property BusinessAddress in the schema.
                /// </summary>
                get: function () {
                    return this._BusinessAddress;
                },
                set: function (value) {
                    if (this._BusinessAddress) {
                        this._BusinessAddress.removeChangedListener(this._BusinessAddressChangedListener);
                    }
                    if (value !== this._BusinessAddress) {
                        this._BusinessAddressChanged = true;
                        this.changed = true;
                    }
                    if (this._BusinessAddress) {
                        this._BusinessAddress.addChangedListener(this._BusinessAddressChangedListener);
                    }
                    this._BusinessAddress = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "businessAddressChanged", {
                get: function () {
                    return this._BusinessAddressChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "otherAddress", {
                /// <summary>
                /// There are no comments for Property OtherAddress in the schema.
                /// </summary>
                get: function () {
                    return this._OtherAddress;
                },
                set: function (value) {
                    if (this._OtherAddress) {
                        this._OtherAddress.removeChangedListener(this._OtherAddressChangedListener);
                    }
                    if (value !== this._OtherAddress) {
                        this._OtherAddressChanged = true;
                        this.changed = true;
                    }
                    if (this._OtherAddress) {
                        this._OtherAddress.addChangedListener(this._OtherAddressChangedListener);
                    }
                    this._OtherAddress = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "otherAddressChanged", {
                get: function () {
                    return this._OtherAddressChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "yomiCompanyName", {
                /// <summary>
                /// There are no comments for Property YomiCompanyName in the schema.
                /// </summary>
                get: function () {
                    return this._YomiCompanyName;
                },
                set: function (value) {
                    if (value !== this._YomiCompanyName) {
                        this._YomiCompanyNameChanged = true;
                        this.changed = true;
                    }
                    this._YomiCompanyName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "yomiCompanyNameChanged", {
                get: function () {
                    return this._YomiCompanyNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "yomiGivenName", {
                /// <summary>
                /// There are no comments for Property YomiGivenName in the schema.
                /// </summary>
                get: function () {
                    return this._YomiGivenName;
                },
                set: function (value) {
                    if (value !== this._YomiGivenName) {
                        this._YomiGivenNameChanged = true;
                        this.changed = true;
                    }
                    this._YomiGivenName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "yomiGivenNameChanged", {
                get: function () {
                    return this._YomiGivenNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Contact.prototype, "yomiSurname", {
                /// <summary>
                /// There are no comments for Property YomiSurname in the schema.
                /// </summary>
                get: function () {
                    return this._YomiSurname;
                },
                set: function (value) {
                    if (value !== this._YomiSurname) {
                        this._YomiSurnameChanged = true;
                        this.changed = true;
                    }
                    this._YomiSurname = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Contact.prototype, "yomiSurnameChanged", {
                get: function () {
                    return this._YomiSurnameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Contact.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Contact.parseContact(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Contact.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Contact.parseContact = function (context, path, data) {
                if (!data)
                    return null;

                return new Contact(context, path, data);
            };

            Contact.parseContacts = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Contact.parseContact(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Contact.prototype.getRequestBody = function () {
                return {
                    ParentFolderId: (this.parentFolderIdChanged && this.parentFolderId) ? this.parentFolderId : undefined,
                    Birthday: (this.birthdayChanged && this.birthday) ? this.birthday.toString() : undefined,
                    FileAs: (this.fileAsChanged && this.fileAs) ? this.fileAs : undefined,
                    DisplayName: (this.displayNameChanged && this.displayName) ? this.displayName : undefined,
                    GivenName: (this.givenNameChanged && this.givenName) ? this.givenName : undefined,
                    Initials: (this.initialsChanged && this.initials) ? this.initials : undefined,
                    MiddleName: (this.middleNameChanged && this.middleName) ? this.middleName : undefined,
                    NickName: (this.nickNameChanged && this.nickName) ? this.nickName : undefined,
                    Surname: (this.surnameChanged && this.surname) ? this.surname : undefined,
                    Title: (this.titleChanged && this.title) ? this.title : undefined,
                    Generation: (this.generationChanged && this.generation) ? this.generation : undefined,
                    EmailAddresses: (this.emailAddressesChanged) ? (function (EmailAddresses) {
                        if (!EmailAddresses) {
                            return undefined;
                        }
                        var converted = [];
                        EmailAddresses.forEach(function (value, index, array) {
                            converted.push(value.getRequestBody());
                        });
                        return converted;
                    })(this.emailAddresses) : undefined,
                    ImAddresses: (this.imAddressesChanged && this.imAddresses) ? this.imAddresses : undefined,
                    JobTitle: (this.jobTitleChanged && this.jobTitle) ? this.jobTitle : undefined,
                    CompanyName: (this.companyNameChanged && this.companyName) ? this.companyName : undefined,
                    Department: (this.departmentChanged && this.department) ? this.department : undefined,
                    OfficeLocation: (this.officeLocationChanged && this.officeLocation) ? this.officeLocation : undefined,
                    Profession: (this.professionChanged && this.profession) ? this.profession : undefined,
                    BusinessHomePage: (this.businessHomePageChanged && this.businessHomePage) ? this.businessHomePage : undefined,
                    AssistantName: (this.assistantNameChanged && this.assistantName) ? this.assistantName : undefined,
                    Manager: (this.managerChanged && this.manager) ? this.manager : undefined,
                    HomePhones: (this.homePhonesChanged && this.homePhones) ? this.homePhones : undefined,
                    BusinessPhones: (this.businessPhonesChanged && this.businessPhones) ? this.businessPhones : undefined,
                    MobilePhone1: (this.mobilePhone1Changed && this.mobilePhone1) ? this.mobilePhone1 : undefined,
                    HomeAddress: (this.homeAddressChanged && this.homeAddress) ? this.homeAddress.getRequestBody() : undefined,
                    BusinessAddress: (this.businessAddressChanged && this.businessAddress) ? this.businessAddress.getRequestBody() : undefined,
                    OtherAddress: (this.otherAddressChanged && this.otherAddress) ? this.otherAddress.getRequestBody() : undefined,
                    YomiCompanyName: (this.yomiCompanyNameChanged && this.yomiCompanyName) ? this.yomiCompanyName : undefined,
                    YomiGivenName: (this.yomiGivenNameChanged && this.yomiGivenName) ? this.yomiGivenName : undefined,
                    YomiSurname: (this.yomiSurnameChanged && this.yomiSurname) ? this.yomiSurname : undefined,
                    ChangeKey: (this.changeKeyChanged && this.changeKey) ? this.changeKey : undefined,
                    Categories: (this.categoriesChanged && this.categories) ? this.categories : undefined,
                    DateTimeCreated: (this.dateTimeCreatedChanged && this.dateTimeCreated) ? this.dateTimeCreated.toString() : undefined,
                    DateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Contact;
        })(Item);
        OutlookServices.Contact = Contact;

        /// <summary>
        /// There are no comments for ContactFolder in the schema.
        /// </summary>
        var ContactFolderFetcher = (function (_super) {
            __extends(ContactFolderFetcher, _super);
            function ContactFolderFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(ContactFolderFetcher.prototype, "contacts", {
                get: function () {
                    if (this._Contacts === undefined) {
                        this._Contacts = new Microsoft.OutlookServices.Contacts(this.context, this.getPath('Contacts'));
                    }
                    return this._Contacts;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(ContactFolderFetcher.prototype, "childFolders", {
                get: function () {
                    if (this._ChildFolders === undefined) {
                        this._ChildFolders = new Microsoft.OutlookServices.ContactFolders(this.context, this.getPath('ChildFolders'));
                    }
                    return this._ChildFolders;
                },
                enumerable: true,
                configurable: true
            });

            ContactFolderFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(ContactFolder.parseContactFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return ContactFolderFetcher;
        })(EntityFetcher);
        OutlookServices.ContactFolderFetcher = ContactFolderFetcher;

        /// <summary>
        /// There are no comments for ContactFolder in the schema.
        /// </summary>
        var ContactFolder = (function (_super) {
            __extends(ContactFolder, _super);
            function ContactFolder(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.OutlookServices.ContactFolder';
                this._ParentFolderIdChanged = false;
                this._DisplayNameChanged = false;

                if (!data) {
                    return;
                }

                this._ParentFolderId = data.ParentFolderId;
                this._DisplayName = data.DisplayName;
            }
            Object.defineProperty(ContactFolder.prototype, "parentFolderId", {
                /// <summary>
                /// There are no comments for Property ParentFolderId in the schema.
                /// </summary>
                get: function () {
                    return this._ParentFolderId;
                },
                set: function (value) {
                    if (value !== this._ParentFolderId) {
                        this._ParentFolderIdChanged = true;
                        this.changed = true;
                    }
                    this._ParentFolderId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ContactFolder.prototype, "parentFolderIdChanged", {
                get: function () {
                    return this._ParentFolderIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(ContactFolder.prototype, "displayName", {
                /// <summary>
                /// There are no comments for Property DisplayName in the schema.
                /// </summary>
                get: function () {
                    return this._DisplayName;
                },
                set: function (value) {
                    if (value !== this._DisplayName) {
                        this._DisplayNameChanged = true;
                        this.changed = true;
                    }
                    this._DisplayName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ContactFolder.prototype, "displayNameChanged", {
                get: function () {
                    return this._DisplayNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(ContactFolder.prototype, "contacts", {
                get: function () {
                    if (this._Contacts === undefined) {
                        this._Contacts = new Microsoft.OutlookServices.Contacts(this.context, this.getPath('Contacts'));
                    }
                    return this._Contacts;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(ContactFolder.prototype, "childFolders", {
                get: function () {
                    if (this._ChildFolders === undefined) {
                        this._ChildFolders = new Microsoft.OutlookServices.ContactFolders(this.context, this.getPath('ChildFolders'));
                    }
                    return this._ChildFolders;
                },
                enumerable: true,
                configurable: true
            });

            ContactFolder.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(ContactFolder.parseContactFolder(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            ContactFolder.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            ContactFolder.parseContactFolder = function (context, path, data) {
                if (!data)
                    return null;

                return new ContactFolder(context, path, data);
            };

            ContactFolder.parseContactFolders = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(ContactFolder.parseContactFolder(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            ContactFolder.prototype.getRequestBody = function () {
                return {
                    ParentFolderId: (this.parentFolderIdChanged && this.parentFolderId) ? this.parentFolderId : undefined,
                    DisplayName: (this.displayNameChanged && this.displayName) ? this.displayName : undefined,
                    Id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return ContactFolder;
        })(Entity);
        OutlookServices.ContactFolder = ContactFolder;

        /// <summary>
        /// There are no comments for DayOfWeek in the schema.
        /// </summary>
        (function (DayOfWeek) {
            DayOfWeek[DayOfWeek["Sunday"] = 0] = "Sunday";
            DayOfWeek[DayOfWeek["Monday"] = 1] = "Monday";
            DayOfWeek[DayOfWeek["Tuesday"] = 2] = "Tuesday";
            DayOfWeek[DayOfWeek["Wednesday"] = 3] = "Wednesday";
            DayOfWeek[DayOfWeek["Thursday"] = 4] = "Thursday";
            DayOfWeek[DayOfWeek["Friday"] = 5] = "Friday";
            DayOfWeek[DayOfWeek["Saturday"] = 6] = "Saturday";
        })(OutlookServices.DayOfWeek || (OutlookServices.DayOfWeek = {}));
        var DayOfWeek = OutlookServices.DayOfWeek;

        /// <summary>
        /// There are no comments for BodyType in the schema.
        /// </summary>
        (function (BodyType) {
            BodyType[BodyType["Text"] = 0] = "Text";
            BodyType[BodyType["HTML"] = 1] = "HTML";
        })(OutlookServices.BodyType || (OutlookServices.BodyType = {}));
        var BodyType = OutlookServices.BodyType;

        /// <summary>
        /// There are no comments for Importance in the schema.
        /// </summary>
        (function (Importance) {
            Importance[Importance["Low"] = 0] = "Low";
            Importance[Importance["Normal"] = 1] = "Normal";
            Importance[Importance["High"] = 2] = "High";
        })(OutlookServices.Importance || (OutlookServices.Importance = {}));
        var Importance = OutlookServices.Importance;

        /// <summary>
        /// There are no comments for AttendeeType in the schema.
        /// </summary>
        (function (AttendeeType) {
            AttendeeType[AttendeeType["Required"] = 0] = "Required";
            AttendeeType[AttendeeType["Optional"] = 1] = "Optional";
            AttendeeType[AttendeeType["Resource"] = 2] = "Resource";
        })(OutlookServices.AttendeeType || (OutlookServices.AttendeeType = {}));
        var AttendeeType = OutlookServices.AttendeeType;

        /// <summary>
        /// There are no comments for ResponseType in the schema.
        /// </summary>
        (function (ResponseType) {
            ResponseType[ResponseType["None"] = 0] = "None";
            ResponseType[ResponseType["Organizer"] = 1] = "Organizer";
            ResponseType[ResponseType["TentativelyAccepted"] = 2] = "TentativelyAccepted";
            ResponseType[ResponseType["Accepted"] = 3] = "Accepted";
            ResponseType[ResponseType["Declined"] = 4] = "Declined";
            ResponseType[ResponseType["NotResponded"] = 5] = "NotResponded";
        })(OutlookServices.ResponseType || (OutlookServices.ResponseType = {}));
        var ResponseType = OutlookServices.ResponseType;

        /// <summary>
        /// There are no comments for EventType in the schema.
        /// </summary>
        (function (EventType) {
            EventType[EventType["SingleInstance"] = 0] = "SingleInstance";
            EventType[EventType["Occurrence"] = 1] = "Occurrence";
            EventType[EventType["Exception"] = 2] = "Exception";
            EventType[EventType["SeriesMaster"] = 3] = "SeriesMaster";
        })(OutlookServices.EventType || (OutlookServices.EventType = {}));
        var EventType = OutlookServices.EventType;

        /// <summary>
        /// There are no comments for FreeBusyStatus in the schema.
        /// </summary>
        (function (FreeBusyStatus) {
            FreeBusyStatus[FreeBusyStatus["Free"] = 0] = "Free";
            FreeBusyStatus[FreeBusyStatus["Tentative"] = 1] = "Tentative";
            FreeBusyStatus[FreeBusyStatus["Busy"] = 2] = "Busy";
            FreeBusyStatus[FreeBusyStatus["Oof"] = 3] = "Oof";
            FreeBusyStatus[FreeBusyStatus["WorkingElsewhere"] = 4] = "WorkingElsewhere";
            FreeBusyStatus[FreeBusyStatus["Unknown"] = -1] = "Unknown";
        })(OutlookServices.FreeBusyStatus || (OutlookServices.FreeBusyStatus = {}));
        var FreeBusyStatus = OutlookServices.FreeBusyStatus;

        /// <summary>
        /// There are no comments for MeetingMessageType in the schema.
        /// </summary>
        (function (MeetingMessageType) {
            MeetingMessageType[MeetingMessageType["None"] = 0] = "None";
            MeetingMessageType[MeetingMessageType["MeetingRequest"] = 1] = "MeetingRequest";
            MeetingMessageType[MeetingMessageType["MeetingCancelled"] = 2] = "MeetingCancelled";
            MeetingMessageType[MeetingMessageType["MeetingAccepted"] = 3] = "MeetingAccepted";
            MeetingMessageType[MeetingMessageType["MeetingTenativelyAccepted"] = 4] = "MeetingTenativelyAccepted";
            MeetingMessageType[MeetingMessageType["MeetingDeclined"] = 5] = "MeetingDeclined";
        })(OutlookServices.MeetingMessageType || (OutlookServices.MeetingMessageType = {}));
        var MeetingMessageType = OutlookServices.MeetingMessageType;

        /// <summary>
        /// There are no comments for RecurrencePatternType in the schema.
        /// </summary>
        (function (RecurrencePatternType) {
            RecurrencePatternType[RecurrencePatternType["Daily"] = 0] = "Daily";
            RecurrencePatternType[RecurrencePatternType["Weekly"] = 1] = "Weekly";
            RecurrencePatternType[RecurrencePatternType["AbsoluteMonthly"] = 2] = "AbsoluteMonthly";
            RecurrencePatternType[RecurrencePatternType["RelativeMonthly"] = 3] = "RelativeMonthly";
            RecurrencePatternType[RecurrencePatternType["AbsoluteYearly"] = 4] = "AbsoluteYearly";
            RecurrencePatternType[RecurrencePatternType["RelativeYearly"] = 5] = "RelativeYearly";
        })(OutlookServices.RecurrencePatternType || (OutlookServices.RecurrencePatternType = {}));
        var RecurrencePatternType = OutlookServices.RecurrencePatternType;

        /// <summary>
        /// There are no comments for RecurrenceRangeType in the schema.
        /// </summary>
        (function (RecurrenceRangeType) {
            RecurrenceRangeType[RecurrenceRangeType["EndDate"] = 0] = "EndDate";
            RecurrenceRangeType[RecurrenceRangeType["NoEnd"] = 1] = "NoEnd";
            RecurrenceRangeType[RecurrenceRangeType["Numbered"] = 2] = "Numbered";
        })(OutlookServices.RecurrenceRangeType || (OutlookServices.RecurrenceRangeType = {}));
        var RecurrenceRangeType = OutlookServices.RecurrenceRangeType;

        /// <summary>
        /// There are no comments for WeekIndex in the schema.
        /// </summary>
        (function (WeekIndex) {
            WeekIndex[WeekIndex["First"] = 0] = "First";
            WeekIndex[WeekIndex["Second"] = 1] = "Second";
            WeekIndex[WeekIndex["Third"] = 2] = "Third";
            WeekIndex[WeekIndex["Fourth"] = 3] = "Fourth";
            WeekIndex[WeekIndex["Last"] = 4] = "Last";
        })(OutlookServices.WeekIndex || (OutlookServices.WeekIndex = {}));
        var WeekIndex = OutlookServices.WeekIndex;
        var Users = (function (_super) {
            __extends(Users, _super);
            function Users(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return User.parseUsers(context, pathFn, data.value);
                };
            }
            Users.prototype.getUser = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new UserFetcher(this.context, path);
                return fetcher;
            };

            Users.prototype.getUsers = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            Users.prototype.addUser = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(User.parseUser(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            return Users;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.Users = Users;
        var Folders = (function (_super) {
            __extends(Folders, _super);
            function Folders(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return Folder.parseFolders(context, pathFn, data.value);
                };
            }
            Folders.prototype.getFolder = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new FolderFetcher(this.context, path);
                return fetcher;
            };

            Folders.prototype.getFolders = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            Folders.prototype.addFolder = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            return Folders;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.Folders = Folders;
        var Messages = (function (_super) {
            __extends(Messages, _super);
            function Messages(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return Message.parseMessages(context, pathFn, data.value);
                };
            }
            Messages.prototype.getMessage = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new MessageFetcher(this.context, path);
                return fetcher;
            };

            Messages.prototype.getMessages = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            Messages.prototype.addMessage = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(Message.parseMessage(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            return Messages;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.Messages = Messages;
        var Calendars = (function (_super) {
            __extends(Calendars, _super);
            function Calendars(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return Calendar.parseCalendars(context, pathFn, data.value);
                };
            }
            Calendars.prototype.getCalendar = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new CalendarFetcher(this.context, path);
                return fetcher;
            };

            Calendars.prototype.getCalendars = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            Calendars.prototype.addCalendar = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(Calendar.parseCalendar(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            return Calendars;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.Calendars = Calendars;
        var CalendarGroups = (function (_super) {
            __extends(CalendarGroups, _super);
            function CalendarGroups(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return CalendarGroup.parseCalendarGroups(context, pathFn, data.value);
                };
            }
            CalendarGroups.prototype.getCalendarGroup = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new CalendarGroupFetcher(this.context, path);
                return fetcher;
            };

            CalendarGroups.prototype.getCalendarGroups = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            CalendarGroups.prototype.addCalendarGroup = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(CalendarGroup.parseCalendarGroup(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            return CalendarGroups;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.CalendarGroups = CalendarGroups;
        var Events = (function (_super) {
            __extends(Events, _super);
            function Events(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return Event.parseEvents(context, pathFn, data.value);
                };
            }
            Events.prototype.getEvent = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new EventFetcher(this.context, path);
                return fetcher;
            };

            Events.prototype.getEvents = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            Events.prototype.addEvent = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(Event.parseEvent(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            return Events;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.Events = Events;
        var Contacts = (function (_super) {
            __extends(Contacts, _super);
            function Contacts(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return Contact.parseContacts(context, pathFn, data.value);
                };
            }
            Contacts.prototype.getContact = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new ContactFetcher(this.context, path);
                return fetcher;
            };

            Contacts.prototype.getContacts = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            Contacts.prototype.addContact = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(Contact.parseContact(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            return Contacts;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.Contacts = Contacts;
        var ContactFolders = (function (_super) {
            __extends(ContactFolders, _super);
            function ContactFolders(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return ContactFolder.parseContactFolders(context, pathFn, data.value);
                };
            }
            ContactFolders.prototype.getContactFolder = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new ContactFolderFetcher(this.context, path);
                return fetcher;
            };

            ContactFolders.prototype.getContactFolders = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            ContactFolders.prototype.addContactFolder = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(ContactFolder.parseContactFolder(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            return ContactFolders;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.ContactFolders = ContactFolders;
        var Attachments = (function (_super) {
            __extends(Attachments, _super);
            function Attachments(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return Attachment.parseAttachments(context, pathFn, data.value);
                };
            }
            Attachments.prototype.getAttachment = function (Id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "Id", type: "Edm.String", value: Id }]);
                var fetcher = new AttachmentFetcher(this.context, path);
                return fetcher;
            };

            Attachments.prototype.getAttachments = function () {
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            Attachments.prototype.addAttachment = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.OutlookServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(Attachment.parseAttachment(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            Attachments.prototype.asFileAttachments = function () {
                var parseCollectionFn = (function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return FileAttachment.parseFileAttachments(context, pathFn, data.value);
                }).bind(this);
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path + '/$/Microsoft.OutlookServices.FileAttachment()', parseCollectionFn);
            };
            Attachments.prototype.asItemAttachments = function () {
                var parseCollectionFn = (function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return ItemAttachment.parseItemAttachments(context, pathFn, data.value);
                }).bind(this);
                return new Microsoft.OutlookServices.Extensions.CollectionQuery(this.context, this.path + '/$/Microsoft.OutlookServices.ItemAttachment()', parseCollectionFn);
            };
            return Attachments;
        })(OutlookServices.Extensions.QueryableSet);
        OutlookServices.Attachments = Attachments;
    })(Microsoft.OutlookServices || (Microsoft.OutlookServices = {}));
    var OutlookServices = Microsoft.OutlookServices;
})(Microsoft || (Microsoft = {}));
//# sourceMappingURL=exchange.js.map
﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------
var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var Microsoft;
(function (Microsoft) {
    (function (CoreServices) {
        (function (Extensions) {
            var ObservableBase = (function () {
                function ObservableBase() {
                    this._changedListeners = [];
                }
                Object.defineProperty(ObservableBase.prototype, "changed", {
                    get: function () {
                        return this._changed;
                    },
                    set: function (value) {
                        var _this = this;
                        this._changed = value;
                        this._changedListeners.forEach((function (value, index, array) {
                            try  {
                                value(_this);
                            } catch (e) {
                            }
                        }).bind(this));
                    },
                    enumerable: true,
                    configurable: true
                });


                ObservableBase.prototype.addChangedListener = function (eventFn) {
                    this._changedListeners.push(eventFn);
                };

                ObservableBase.prototype.removeChangedListener = function (eventFn) {
                    var index = this._changedListeners.indexOf(eventFn);
                    if (index >= 0) {
                        this._changedListeners.splice(index, 1);
                    }
                };
                return ObservableBase;
            })();
            Extensions.ObservableBase = ObservableBase;

            var ObservableCollection = (function (_super) {
                __extends(ObservableCollection, _super);
                function ObservableCollection() {
                    var items = [];
                    for (var _i = 0; _i < (arguments.length - 0); _i++) {
                        items[_i] = arguments[_i + 0];
                    }
                    var _this = this;
                    _super.call(this);
                    this._changedListener = (function (changed) {
                        _this.changed = true;
                    }).bind(this);
                    this._array = items;
                }
                ObservableCollection.prototype.item = function (n) {
                    return this._array[n];
                };

                /**
                * Removes the last element from an array and returns it.
                */
                ObservableCollection.prototype.pop = function () {
                    this.changed = true;
                    var result = this._array.pop();
                    result.removeChangedListener(this._changedListener);
                    return result;
                };

                /**
                * Removes the first element from an array and returns it.
                */
                ObservableCollection.prototype.shift = function () {
                    this.changed = true;
                    var result = this._array.shift();
                    result.removeChangedListener(this._changedListener);
                    return result;
                };

                /**
                * Appends new elements to an array, and returns the new length of the array.
                * @param items New elements of the Array.
                */
                ObservableCollection.prototype.push = function () {
                    var _this = this;
                    var items = [];
                    for (var _i = 0; _i < (arguments.length - 0); _i++) {
                        items[_i] = arguments[_i + 0];
                    }
                    items.forEach((function (value, index, array) {
                        try  {
                            value.addChangedListener(_this._changedListener);
                            _this._array.push(value);
                        } catch (e) {
                        }
                    }).bind(this));
                    this.changed = true;
                    return this._array.length;
                };

                /**
                * Removes elements from an array, returning the deleted elements.
                * @param start The zero-based location in the array from which to start removing elements.
                * @param deleteCount The number of elements to remove.
                * @param items Elements to insert into the array in place of the deleted elements.
                */
                ObservableCollection.prototype.splice = function (start, deleteCount) {
                    var _this = this;
                    var result = this._array.splice(start, deleteCount);
                    result.forEach((function (value, index, array) {
                        try  {
                            value.removeChangedListener(_this._changedListener);
                        } catch (e) {
                        }
                    }).bind(this));
                    this.changed = true;
                    return result;
                };

                /**
                * Inserts new elements at the start of an array.
                * @param items  Elements to insert at the start of the Array.
                */
                ObservableCollection.prototype.unshift = function () {
                    var items = [];
                    for (var _i = 0; _i < (arguments.length - 0); _i++) {
                        items[_i] = arguments[_i + 0];
                    }
                    for (var index = items.length - 1; index >= 0; index--) {
                        try  {
                            items[index].addChangedListener(this._changedListener);
                            this._array.unshift(items[index]);
                        } catch (e) {
                        }
                    }
                    this.changed = true;
                    return this._array.length;
                };

                /**
                * Performs the specified action for each element in an array.
                * @param callbackfn  A function that accepts up to three arguments. forEach calls the callbackfn function one time for each element in the array.
                * @param thisArg  An object to which the this keyword can refer in the callbackfn function. If thisArg is omitted, undefined is used as the this value.
                */
                ObservableCollection.prototype.forEach = function (callbackfn, thisArg) {
                    this._array.forEach(callbackfn, thisArg);
                };

                /**
                * Calls a defined callback function on each element of an array, and returns an array that contains the results.
                * @param callbackfn A function that accepts up to three arguments. The map method calls the callbackfn function one time for each element in the array.
                * @param thisArg An object to which the this keyword can refer in the callbackfn function. If thisArg is omitted, undefined is used as the this value.
                */
                ObservableCollection.prototype.map = function (callbackfn, thisArg) {
                    return this._array.map(callbackfn, thisArg);
                };

                /**
                * Returns the elements of an array that meet the condition specified in a callback function.
                * @param callbackfn A function that accepts up to three arguments. The filter method calls the callbackfn function one time for each element in the array.
                * @param thisArg An object to which the this keyword can refer in the callbackfn function. If thisArg is omitted, undefined is used as the this value.
                */
                ObservableCollection.prototype.filter = function (callbackfn, thisArg) {
                    return this._array.filter(callbackfn, thisArg);
                };

                /**
                * Calls the specified callback function for all the elements in an array. The return value of the callback function is the accumulated result, and is provided as an argument in the next call to the callback function.
                * @param callbackfn A function that accepts up to four arguments. The reduce method calls the callbackfn function one time for each element in the array.
                * @param initialValue If initialValue is specified, it is used as the initial value to start the accumulation. The first call to the callbackfn function provides this value as an argument instead of an array value.
                */
                ObservableCollection.prototype.reduce = function (callbackfn, initialValue) {
                    return this._array.reduce(callbackfn, initialValue);
                };

                /**
                * Calls the specified callback function for all the elements in an array, in descending order. The return value of the callback function is the accumulated result, and is provided as an argument in the next call to the callback function.
                * @param callbackfn A function that accepts up to four arguments. The reduceRight method calls the callbackfn function one time for each element in the array.
                * @param initialValue If initialValue is specified, it is used as the initial value to start the accumulation. The first call to the callbackfn function provides this value as an argument instead of an array value.
                */
                ObservableCollection.prototype.reduceRight = function (callbackfn, initialValue) {
                    return this._array.reduceRight(callbackfn, initialValue);
                };

                Object.defineProperty(ObservableCollection.prototype, "length", {
                    /**
                    * Gets or sets the length of the array. This is a number one higher than the highest element defined in an array.
                    */
                    get: function () {
                        return this._array.length;
                    },
                    enumerable: true,
                    configurable: true
                });
                return ObservableCollection;
            })(ObservableBase);
            Extensions.ObservableCollection = ObservableCollection;

            var Request = (function () {
                function Request(requestUri) {
                    this.requestUri = requestUri;
                    this.headers = {};
                    this.disableCache = false;
                }
                return Request;
            })();
            Extensions.Request = Request;

            var DataContext = (function () {
                function DataContext(serviceRootUri, extraQueryParameters, getAccessTokenFn) {
                    this._noCache = Date.now();
                    this.serviceRootUri = serviceRootUri;
                    this.extraQueryParameters = extraQueryParameters;
                    this._getAccessTokenFn = getAccessTokenFn;
                }
                Object.defineProperty(DataContext.prototype, "serviceRootUri", {
                    get: function () {
                        return this._serviceRootUri;
                    },
                    set: function (value) {
                        if (value.lastIndexOf("/") === value.length - 1) {
                            value = value.substring(0, value.length - 1);
                        }

                        this._serviceRootUri = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                Object.defineProperty(DataContext.prototype, "extraQueryParameters", {
                    get: function () {
                        return this._extraQueryParameters;
                    },
                    set: function (value) {
                        this._extraQueryParameters = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                Object.defineProperty(DataContext.prototype, "disableCache", {
                    get: function () {
                        return this._disableCache;
                    },
                    set: function (value) {
                        this._disableCache = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                Object.defineProperty(DataContext.prototype, "disableCacheOverride", {
                    get: function () {
                        return this._disableCacheOverride;
                    },
                    set: function (value) {
                        this._disableCacheOverride = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                DataContext.prototype.ajax = function (request) {
                    var deferred = new Microsoft.Utility.Deferred();

                    var xhr = new XMLHttpRequest(); xhr.responseType = request.responseType;

                    if (!request.method) {
                        request.method = 'GET';
                    }

                    xhr.open(request.method.toUpperCase(), request.requestUri, true);

                    if (request.headers) {
                        for (name in request.headers) {
                            xhr.setRequestHeader(name, request.headers[name]);
                        }
                    }

                    xhr.onreadystatechange = function (e) {
                        if (xhr.readyState == 4) {
                            if (xhr.status >= 200 && xhr.status < 300 || xhr.status === 304) {
                                deferred.resolve(xhr.response);
                            } else {
                                deferred.reject(xhr);
                            }
                        } else {
                            deferred.notify(xhr.readyState);
                        }
                    };

                    if (request.data) {
                        if ((typeof request.data === 'string') || (Buffer.isBuffer(request.data))) {
                            xhr.send(request.data);
                        } else {
                            xhr.send(JSON.stringify(request.data));
                        }
                    } else {
                        xhr.send();
                    }

                    return deferred;
                };

                DataContext.prototype.read = function (path) {
                    return this.request(new Request(this.serviceRootUri + ((this.serviceRootUri.lastIndexOf('/') != this.serviceRootUri.length - 1) ? '/' : '') + path));
                };

                DataContext.prototype.readUrl = function (url) {
                    return this.request(new Request(url));
                };

                DataContext.prototype.request = function (request) {
                    var _this = this;
                    var deferred;

                    this.augmentRequest(request);

                    if (this._getAccessTokenFn) {
                        deferred = new Microsoft.Utility.Deferred();

                        this._getAccessTokenFn().then((function (token) {
                            request.headers["X-ClientService-ClientTag"] = 'Office 365 API Tools, 1.1.0512';
                            request.headers["Authorization"] = 'Bearer ' + token;
                            _this.ajax(request).then(deferred.resolve.bind(deferred), deferred.reject.bind(deferred));
                        }).bind(this), deferred.reject.bind(deferred));
                    } else {
                        deferred = this.ajax(request);
                    }

                    return deferred;
                };

                DataContext.prototype.augmentRequest = function (request) {
                    if (!request.headers) {
                        request.headers = {};
                    }

                    if (!request.headers['Accept']) {
                        request.headers['Accept'] = 'application/json';
                    }

                    if (!request.headers['Content-Type']) {
                        request.headers['Content-Type'] = 'application/json';
                    }

                    if (this.extraQueryParameters) {
                        request.requestUri += (request.requestUri.indexOf('?') >= 0 ? '&' : '?') + this.extraQueryParameters;
                    }

                    if ((!this._disableCacheOverride && request.disableCache) || (this._disableCacheOverride && this._disableCache)) {
                        request.requestUri += (request.requestUri.indexOf('?') >= 0 ? '&' : '?') + '_=' + this._noCache++;
                    }
                };
                return DataContext;
            })();
            Extensions.DataContext = DataContext;

            var PagedCollection = (function () {
                function PagedCollection(context, path, resultFn, data) {
                    this._context = context;
                    this._path = path;
                    this._resultFn = resultFn;
                    this._data = data;
                }
                Object.defineProperty(PagedCollection.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(PagedCollection.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(PagedCollection.prototype, "currentPage", {
                    get: function () {
                        return this._data;
                    },
                    enumerable: true,
                    configurable: true
                });

                PagedCollection.prototype.getNextPage = function () {
                    var _this = this;
                    var deferred = new Microsoft.Utility.Deferred();

                    if (this.path == null) {
                        deferred.resolve(null);
                        return deferred;
                    }

                    var request = new Request(this.path);

                    request.disableCache = true;

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data), nextLink = (parsedData['odata.nextLink'] === undefined) ? ((parsedData['@odata.nextLink'] === undefined) ? ((parsedData['__next'] === undefined) ? null : parsedData['__next']) : parsedData['@odata.nextLink']) : parsedData['odata.nextLink'];

                        deferred.resolve(new PagedCollection(_this.context, nextLink, _this._resultFn, _this._resultFn(_this.context, parsedData)));
                    }).bind(this), deferred.reject.bind(deferred));

                    return deferred;
                };
                return PagedCollection;
            })();
            Extensions.PagedCollection = PagedCollection;

            var CollectionQuery = (function () {
                function CollectionQuery(context, path, resultFn) {
                    this._context = context;
                    this._path = path;
                    this._resultFn = resultFn;
                }
                Object.defineProperty(CollectionQuery.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(CollectionQuery.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                CollectionQuery.prototype.filter = function (filter) {
                    this.addQuery("$filter=" + filter);
                    return this;
                };

                CollectionQuery.prototype.select = function (selection) {
                    if (typeof selection === 'string') {
                        this.addQuery("$select=" + selection);
                    } else if (Array.isArray(selection)) {
                        this.addQuery("$select=" + selection.join(','));
                    } else {
                        throw new Microsoft.Utility.Exception('\'select\' argument must be string or string[].');
                    }
                    return this;
                };

                CollectionQuery.prototype.expand = function (expand) {
                    if (typeof expand === 'string') {
                        this.addQuery("$expand=" + expand);
                    } else if (Array.isArray(expand)) {
                        this.addQuery("$expand=" + expand.join(','));
                    } else {
                        throw new Microsoft.Utility.Exception('\'expand\' argument must be string or string[].');
                    }
                    return this;
                };

                CollectionQuery.prototype.orderBy = function (orderBy) {
                    if (typeof orderBy === 'string') {
                        this.addQuery("$orderby=" + orderBy);
                    } else if (Array.isArray(orderBy)) {
                        this.addQuery("$orderby=" + orderBy.join(','));
                    } else {
                        throw new Microsoft.Utility.Exception('\'orderBy\' argument must be string or string[].');
                    }
                    return this;
                };

                CollectionQuery.prototype.top = function (top) {
                    this.addQuery("$top=" + top);
                    return this;
                };

                CollectionQuery.prototype.skip = function (skip) {
                    this.addQuery("$skip=" + skip);
                    return this;
                };

                CollectionQuery.prototype.addQuery = function (query) {
                    this._query = (this._query ? this._query + "&" : "") + query;
                    return this;
                };

                Object.defineProperty(CollectionQuery.prototype, "query", {
                    get: function () {
                        return this._query;
                    },
                    set: function (value) {
                        this._query = value;
                    },
                    enumerable: true,
                    configurable: true
                });


                CollectionQuery.prototype.fetch = function () {
                    var path = this.path + (this._query ? (this.path.indexOf('?') < 0 ? '?' : '&') + this._query : "");

                    return new Microsoft.CoreServices.Extensions.PagedCollection(this.context, path, this._resultFn).getNextPage();
                };

                CollectionQuery.prototype.fetchAll = function (maxItems) {
                    var path = this.path + (this._query ? (this.path.indexOf('?') < 0 ? '?' : '&') + this._query : ""), pagedItems = new Microsoft.CoreServices.Extensions.PagedCollection(this.context, path, this._resultFn), accumulator = [], deferred = new Microsoft.Utility.Deferred(), recursive = function (nextPagedItems) {
                        if (!nextPagedItems) {
                            deferred.resolve(accumulator);
                        } else {
                            accumulator = accumulator.concat(nextPagedItems.currentPage);

                            if (accumulator.length > maxItems) {
                                accumulator = accumulator.splice(maxItems);
                                deferred.resolve(accumulator);
                            } else {
                                nextPagedItems.getNextPage().then(function (nextPage) {
                                    return recursive(nextPage);
                                }, deferred.reject.bind(deferred));
                            }
                        }
                    };

                    pagedItems.getNextPage().then(function (nextPage) {
                        return recursive(nextPage);
                    }, deferred.reject.bind(deferred));

                    return deferred;
                };
                return CollectionQuery;
            })();
            Extensions.CollectionQuery = CollectionQuery;

            var QueryableSet = (function () {
                function QueryableSet(context, path, entity) {
                    this._context = context;
                    this._path = path;
                    this._entity = entity;
                }
                Object.defineProperty(QueryableSet.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(QueryableSet.prototype, "entity", {
                    get: function () {
                        return this._entity;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(QueryableSet.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                QueryableSet.prototype.getPath = function (prop) {
                    return this._path + '/' + prop;
                };
                return QueryableSet;
            })();
            Extensions.QueryableSet = QueryableSet;

            var RestShallowObjectFetcher = (function () {
                function RestShallowObjectFetcher(context, path) {
                    this._path = path;
                    this._context = context;
                }
                Object.defineProperty(RestShallowObjectFetcher.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(RestShallowObjectFetcher.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                RestShallowObjectFetcher.prototype.getPath = function (prop) {
                    return this._path + '/' + prop;
                };
                return RestShallowObjectFetcher;
            })();
            Extensions.RestShallowObjectFetcher = RestShallowObjectFetcher;

            var ComplexTypeBase = (function (_super) {
                __extends(ComplexTypeBase, _super);
                function ComplexTypeBase() {
                    _super.call(this);
                }
                return ComplexTypeBase;
            })(ObservableBase);
            Extensions.ComplexTypeBase = ComplexTypeBase;

            var EntityBase = (function (_super) {
                __extends(EntityBase, _super);
                function EntityBase(context, path) {
                    _super.call(this);
                    this._path = path;
                    this._context = context;
                }
                Object.defineProperty(EntityBase.prototype, "context", {
                    get: function () {
                        return this._context;
                    },
                    enumerable: true,
                    configurable: true
                });

                Object.defineProperty(EntityBase.prototype, "path", {
                    get: function () {
                        return this._path;
                    },
                    enumerable: true,
                    configurable: true
                });

                EntityBase.prototype.getPath = function (prop) {
                    return this._path + '/' + prop;
                };
                return EntityBase;
            })(ObservableBase);
            Extensions.EntityBase = EntityBase;

            /*
            std
            */
            function isUndefined(v) {
                return typeof v === 'undefined';
            }
            Extensions.isUndefined = isUndefined;
        })(CoreServices.Extensions || (CoreServices.Extensions = {}));
        var Extensions = CoreServices.Extensions;
    })(Microsoft.CoreServices || (Microsoft.CoreServices = {}));
    var CoreServices = Microsoft.CoreServices;
})(Microsoft || (Microsoft = {}));

var Microsoft;
(function (Microsoft) {
    (function (CoreServices) {
        /// <summary>
        /// There are no comments for EntityContainer in the schema.
        /// </summary>
        var SharePointClient = (function () {
            function SharePointClient(serviceRootUri, getAccessTokenFn) {
                this._context = new Microsoft.CoreServices.Extensions.DataContext(serviceRootUri, undefined, getAccessTokenFn);
            }
            Object.defineProperty(SharePointClient.prototype, "context", {
                get: function () {
                    return this._context;
                },
                enumerable: true,
                configurable: true
            });

            SharePointClient.prototype.getPath = function (prop) {
                return this.context.serviceRootUri + '/' + prop;
            };

            Object.defineProperty(SharePointClient.prototype, "files", {
                get: function () {
                    if (this._files === undefined) {
                        this._files = new Microsoft.FileServices.Items(this.context, this.getPath('files'));
                    }
                    return this._files;
                },
                enumerable: true,
                configurable: true
            });

            /// <summary>
            /// There are no comments for files in the schema.
            /// </summary>
            SharePointClient.prototype.addTofiles = function (item) {
                this.files.addItem(item);
            };

            Object.defineProperty(SharePointClient.prototype, "drive", {
                get: function () {
                    if (this._drive === undefined) {
                        this._drive = new Microsoft.FileServices.DriveFetcher(this.context, this.getPath("drive"));
                    }
                    return this._drive;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(SharePointClient.prototype, "me", {
                get: function () {
                    if (this._me === undefined) {
                        this._me = new CurrentUserRequestContextFetcher(this.context, this.getPath("me"));
                    }
                    return this._me;
                },
                enumerable: true,
                configurable: true
            });
            return SharePointClient;
        })();
        CoreServices.SharePointClient = SharePointClient;

        /// <summary>
        /// There are no comments for CurrentUserRequestContext in the schema.
        /// </summary>
        var CurrentUserRequestContextFetcher = (function (_super) {
            __extends(CurrentUserRequestContextFetcher, _super);
            function CurrentUserRequestContextFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(CurrentUserRequestContextFetcher.prototype, "drive", {
                /// <summary>
                /// There are no comments for Query Property drive in the schema.
                /// </summary>
                get: function () {
                    if (this._drive === undefined) {
                        this._drive = new Microsoft.FileServices.DriveFetcher(this.context, this.getPath("drive"));
                    }
                    return this._drive;
                },
                enumerable: true,
                configurable: true
            });

            CurrentUserRequestContextFetcher.prototype.update_drive = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("$links/drive"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Object.defineProperty(CurrentUserRequestContextFetcher.prototype, "files", {
                get: function () {
                    if (this._files === undefined) {
                        this._files = new Microsoft.FileServices.Items(this.context, this.getPath('files'));
                    }
                    return this._files;
                },
                enumerable: true,
                configurable: true
            });

            CurrentUserRequestContextFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(CurrentUserRequestContext.parseCurrentUserRequestContext(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return CurrentUserRequestContextFetcher;
        })(CoreServices.Extensions.RestShallowObjectFetcher);
        CoreServices.CurrentUserRequestContextFetcher = CurrentUserRequestContextFetcher;

        /// <summary>
        /// There are no comments for CurrentUserRequestContext in the schema.
        /// </summary>
        var CurrentUserRequestContext = (function (_super) {
            __extends(CurrentUserRequestContext, _super);
            function CurrentUserRequestContext(context, path, data) {
                _super.call(this, context, path);
                this._odataType = 'Microsoft.CoreServices.CurrentUserRequestContext';
                this._idChanged = false;

                if (!data) {
                    return;
                }

                this._id = data.id;
            }
            Object.defineProperty(CurrentUserRequestContext.prototype, "id", {
                /// <summary>
                /// There are no comments for Property id in the schema.
                /// </summary>
                get: function () {
                    return this._id;
                },
                set: function (value) {
                    if (value !== this._id) {
                        this._idChanged = true;
                        this.changed = true;
                    }
                    this._id = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(CurrentUserRequestContext.prototype, "idChanged", {
                get: function () {
                    return this._idChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(CurrentUserRequestContext.prototype, "drive", {
                /// <summary>
                /// There are no comments for Query Property drive in the schema.
                /// </summary>
                get: function () {
                    if (this._drive === undefined) {
                        this._drive = new Microsoft.FileServices.DriveFetcher(this.context, this.getPath("drive"));
                    }
                    return this._drive;
                },
                enumerable: true,
                configurable: true
            });

            CurrentUserRequestContext.prototype.update_drive = function (value) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("$links/drive"));

                request.method = 'PUT';
                request.data = JSON.stringify({ url: value.path });

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Object.defineProperty(CurrentUserRequestContext.prototype, "files", {
                get: function () {
                    if (this._files === undefined) {
                        this._files = new Microsoft.FileServices.Items(this.context, this.getPath('files'));
                    }
                    return this._files;
                },
                enumerable: true,
                configurable: true
            });

            CurrentUserRequestContext.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(CurrentUserRequestContext.parseCurrentUserRequestContext(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            CurrentUserRequestContext.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            CurrentUserRequestContext.parseCurrentUserRequestContext = function (context, path, data) {
                if (!data)
                    return null;

                return new CurrentUserRequestContext(context, path, data);
            };

            CurrentUserRequestContext.parseCurrentUserRequestContexts = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(CurrentUserRequestContext.parseCurrentUserRequestContext(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            CurrentUserRequestContext.prototype.getRequestBody = function () {
                return {
                    id: (this.idChanged && this.id) ? this.id : undefined,
                    '@odata.type': this._odataType
                };
            };
            return CurrentUserRequestContext;
        })(CoreServices.Extensions.EntityBase);
        CoreServices.CurrentUserRequestContext = CurrentUserRequestContext;
    })(Microsoft.CoreServices || (Microsoft.CoreServices = {}));
    var CoreServices = Microsoft.CoreServices;
})(Microsoft || (Microsoft = {}));

var Microsoft;
(function (Microsoft) {
    (function (FileServices) {
        /// <summary>
        /// There are no comments for DriveQuota in the schema.
        /// </summary>
        var DriveQuota = (function (_super) {
            __extends(DriveQuota, _super);
            function DriveQuota(data) {
                _super.call(this);
                this._odataType = 'Microsoft.FileServices.DriveQuota';
                this._deletedChanged = false;
                this._remainingChanged = false;
                this._stateChanged = false;
                this._totalChanged = false;

                if (!data) {
                    return;
                }

                this._deleted = data.deleted;
                this._remaining = data.remaining;
                this._state = data.state;
                this._total = data.total;
            }
            Object.defineProperty(DriveQuota.prototype, "deleted", {
                /// <summary>
                /// There are no comments for Property deleted in the schema.
                /// </summary>
                get: function () {
                    return this._deleted;
                },
                set: function (value) {
                    if (value !== this._deleted) {
                        this._deletedChanged = true;
                        this.changed = true;
                    }
                    this._deleted = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(DriveQuota.prototype, "deletedChanged", {
                get: function () {
                    return this._deletedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(DriveQuota.prototype, "remaining", {
                /// <summary>
                /// There are no comments for Property remaining in the schema.
                /// </summary>
                get: function () {
                    return this._remaining;
                },
                set: function (value) {
                    if (value !== this._remaining) {
                        this._remainingChanged = true;
                        this.changed = true;
                    }
                    this._remaining = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(DriveQuota.prototype, "remainingChanged", {
                get: function () {
                    return this._remainingChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(DriveQuota.prototype, "state", {
                /// <summary>
                /// There are no comments for Property state in the schema.
                /// </summary>
                get: function () {
                    return this._state;
                },
                set: function (value) {
                    if (value !== this._state) {
                        this._stateChanged = true;
                        this.changed = true;
                    }
                    this._state = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(DriveQuota.prototype, "stateChanged", {
                get: function () {
                    return this._stateChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(DriveQuota.prototype, "total", {
                /// <summary>
                /// There are no comments for Property total in the schema.
                /// </summary>
                get: function () {
                    return this._total;
                },
                set: function (value) {
                    if (value !== this._total) {
                        this._totalChanged = true;
                        this.changed = true;
                    }
                    this._total = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(DriveQuota.prototype, "totalChanged", {
                get: function () {
                    return this._totalChanged;
                },
                enumerable: true,
                configurable: true
            });

            DriveQuota.parseDriveQuota = function (data) {
                if (!data)
                    return null;

                return new DriveQuota(data);
            };

            DriveQuota.parseDriveQuotas = function (data) {
                var results = new Microsoft.CoreServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(DriveQuota.parseDriveQuota(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            DriveQuota.prototype.getRequestBody = function () {
                return {
                    deleted: (this.deletedChanged && this.deleted) ? this.deleted : undefined,
                    remaining: (this.remainingChanged && this.remaining) ? this.remaining : undefined,
                    state: (this.stateChanged && this.state) ? this.state : undefined,
                    total: (this.totalChanged && this.total) ? this.total : undefined,
                    '@odata.type': this._odataType
                };
            };
            return DriveQuota;
        })(Microsoft.CoreServices.Extensions.ComplexTypeBase);
        FileServices.DriveQuota = DriveQuota;

        /// <summary>
        /// There are no comments for IdentitySet in the schema.
        /// </summary>
        var IdentitySet = (function (_super) {
            __extends(IdentitySet, _super);
            function IdentitySet(data) {
                var _this = this;
                _super.call(this);
                this._odataType = 'Microsoft.FileServices.IdentitySet';
                this._applicationChanged = false;
                this._applicationChangedListener = (function (value) {
                    _this._applicationChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._userChanged = false;
                this._userChangedListener = (function (value) {
                    _this._userChanged = true;
                    _this.changed = true;
                }).bind(this);

                if (!data) {
                    return;
                }

                this._application = Identity.parseIdentity(data.application);
                if (this._application) {
                    this._application.addChangedListener(this._applicationChangedListener);
                }
                this._user = Identity.parseIdentity(data.user);
                if (this._user) {
                    this._user.addChangedListener(this._userChangedListener);
                }
            }
            Object.defineProperty(IdentitySet.prototype, "application", {
                /// <summary>
                /// There are no comments for Property application in the schema.
                /// </summary>
                get: function () {
                    return this._application;
                },
                set: function (value) {
                    if (this._application) {
                        this._application.removeChangedListener(this._applicationChangedListener);
                    }
                    if (value !== this._application) {
                        this._applicationChanged = true;
                        this.changed = true;
                    }
                    if (this._application) {
                        this._application.addChangedListener(this._applicationChangedListener);
                    }
                    this._application = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(IdentitySet.prototype, "applicationChanged", {
                get: function () {
                    return this._applicationChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(IdentitySet.prototype, "user", {
                /// <summary>
                /// There are no comments for Property user in the schema.
                /// </summary>
                get: function () {
                    return this._user;
                },
                set: function (value) {
                    if (this._user) {
                        this._user.removeChangedListener(this._userChangedListener);
                    }
                    if (value !== this._user) {
                        this._userChanged = true;
                        this.changed = true;
                    }
                    if (this._user) {
                        this._user.addChangedListener(this._userChangedListener);
                    }
                    this._user = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(IdentitySet.prototype, "userChanged", {
                get: function () {
                    return this._userChanged;
                },
                enumerable: true,
                configurable: true
            });

            IdentitySet.parseIdentitySet = function (data) {
                if (!data)
                    return null;

                return new IdentitySet(data);
            };

            IdentitySet.parseIdentitySets = function (data) {
                var results = new Microsoft.CoreServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(IdentitySet.parseIdentitySet(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            IdentitySet.prototype.getRequestBody = function () {
                return {
                    application: (this.applicationChanged && this.application) ? this.application.getRequestBody() : undefined,
                    user: (this.userChanged && this.user) ? this.user.getRequestBody() : undefined,
                    '@odata.type': this._odataType
                };
            };
            return IdentitySet;
        })(Microsoft.CoreServices.Extensions.ComplexTypeBase);
        FileServices.IdentitySet = IdentitySet;

        /// <summary>
        /// There are no comments for Identity in the schema.
        /// </summary>
        var Identity = (function (_super) {
            __extends(Identity, _super);
            function Identity(data) {
                _super.call(this);
                this._odataType = 'Microsoft.FileServices.Identity';
                this._idChanged = false;
                this._displayNameChanged = false;

                if (!data) {
                    return;
                }

                this._id = data.id;
                this._displayName = data.displayName;
            }
            Object.defineProperty(Identity.prototype, "id", {
                /// <summary>
                /// There are no comments for Property id in the schema.
                /// </summary>
                get: function () {
                    return this._id;
                },
                set: function (value) {
                    if (value !== this._id) {
                        this._idChanged = true;
                        this.changed = true;
                    }
                    this._id = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Identity.prototype, "idChanged", {
                get: function () {
                    return this._idChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Identity.prototype, "displayName", {
                /// <summary>
                /// There are no comments for Property displayName in the schema.
                /// </summary>
                get: function () {
                    return this._displayName;
                },
                set: function (value) {
                    if (value !== this._displayName) {
                        this._displayNameChanged = true;
                        this.changed = true;
                    }
                    this._displayName = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Identity.prototype, "displayNameChanged", {
                get: function () {
                    return this._displayNameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Identity.parseIdentity = function (data) {
                if (!data)
                    return null;

                return new Identity(data);
            };

            Identity.parseIdentities = function (data) {
                var results = new Microsoft.CoreServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Identity.parseIdentity(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            Identity.prototype.getRequestBody = function () {
                return {
                    id: (this.idChanged && this.id) ? this.id : undefined,
                    displayName: (this.displayNameChanged && this.displayName) ? this.displayName : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Identity;
        })(Microsoft.CoreServices.Extensions.ComplexTypeBase);
        FileServices.Identity = Identity;

        /// <summary>
        /// There are no comments for ItemReference in the schema.
        /// </summary>
        var ItemReference = (function (_super) {
            __extends(ItemReference, _super);
            function ItemReference(data) {
                _super.call(this);
                this._odataType = 'Microsoft.FileServices.ItemReference';
                this._driveIdChanged = false;
                this._idChanged = false;
                this._pathChanged = false;

                if (!data) {
                    return;
                }

                this._driveId = data.driveId;
                this._id = data.id;
                this._path = data.path;
            }
            Object.defineProperty(ItemReference.prototype, "driveId", {
                /// <summary>
                /// There are no comments for Property driveId in the schema.
                /// </summary>
                get: function () {
                    return this._driveId;
                },
                set: function (value) {
                    if (value !== this._driveId) {
                        this._driveIdChanged = true;
                        this.changed = true;
                    }
                    this._driveId = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ItemReference.prototype, "driveIdChanged", {
                get: function () {
                    return this._driveIdChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(ItemReference.prototype, "id", {
                /// <summary>
                /// There are no comments for Property id in the schema.
                /// </summary>
                get: function () {
                    return this._id;
                },
                set: function (value) {
                    if (value !== this._id) {
                        this._idChanged = true;
                        this.changed = true;
                    }
                    this._id = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ItemReference.prototype, "idChanged", {
                get: function () {
                    return this._idChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(ItemReference.prototype, "path", {
                /// <summary>
                /// There are no comments for Property path in the schema.
                /// </summary>
                get: function () {
                    return this._path;
                },
                set: function (value) {
                    if (value !== this._path) {
                        this._pathChanged = true;
                        this.changed = true;
                    }
                    this._path = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(ItemReference.prototype, "pathChanged", {
                get: function () {
                    return this._pathChanged;
                },
                enumerable: true,
                configurable: true
            });

            ItemReference.parseItemReference = function (data) {
                if (!data)
                    return null;

                return new ItemReference(data);
            };

            ItemReference.parseItemReferences = function (data) {
                var results = new Microsoft.CoreServices.Extensions.ObservableCollection();

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(ItemReference.parseItemReference(data[i]));
                    }
                }

                results.changed = false;

                return results;
            };

            ItemReference.prototype.getRequestBody = function () {
                return {
                    driveId: (this.driveIdChanged && this.driveId) ? this.driveId : undefined,
                    id: (this.idChanged && this.id) ? this.id : undefined,
                    path: (this.pathChanged && this.path) ? this.path : undefined,
                    '@odata.type': this._odataType
                };
            };
            return ItemReference;
        })(Microsoft.CoreServices.Extensions.ComplexTypeBase);
        FileServices.ItemReference = ItemReference;

        /// <summary>
        /// There are no comments for Drive in the schema.
        /// </summary>
        var DriveFetcher = (function (_super) {
            __extends(DriveFetcher, _super);
            function DriveFetcher(context, path) {
                _super.call(this, context, path);
            }
            DriveFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Drive.parseDrive(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return DriveFetcher;
        })(Microsoft.CoreServices.Extensions.RestShallowObjectFetcher);
        FileServices.DriveFetcher = DriveFetcher;

        /// <summary>
        /// There are no comments for Drive in the schema.
        /// </summary>
        var Drive = (function (_super) {
            __extends(Drive, _super);
            function Drive(context, path, data) {
                var _this = this;
                _super.call(this, context, path);
                this._odataType = 'Microsoft.FileServices.Drive';
                this._idChanged = false;
                this._ownerChanged = false;
                this._ownerChangedListener = (function (value) {
                    _this._ownerChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._quotaChanged = false;
                this._quotaChangedListener = (function (value) {
                    _this._quotaChanged = true;
                    _this.changed = true;
                }).bind(this);

                if (!data) {
                    return;
                }

                this._id = data.id;
                this._owner = Identity.parseIdentity(data.owner);
                if (this._owner) {
                    this._owner.addChangedListener(this._ownerChangedListener);
                }
                this._quota = DriveQuota.parseDriveQuota(data.quota);
                if (this._quota) {
                    this._quota.addChangedListener(this._quotaChangedListener);
                }
            }
            Object.defineProperty(Drive.prototype, "id", {
                /// <summary>
                /// There are no comments for Property id in the schema.
                /// </summary>
                get: function () {
                    return this._id;
                },
                set: function (value) {
                    if (value !== this._id) {
                        this._idChanged = true;
                        this.changed = true;
                    }
                    this._id = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Drive.prototype, "idChanged", {
                get: function () {
                    return this._idChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Drive.prototype, "owner", {
                /// <summary>
                /// There are no comments for Property owner in the schema.
                /// </summary>
                get: function () {
                    return this._owner;
                },
                set: function (value) {
                    if (this._owner) {
                        this._owner.removeChangedListener(this._ownerChangedListener);
                    }
                    if (value !== this._owner) {
                        this._ownerChanged = true;
                        this.changed = true;
                    }
                    if (this._owner) {
                        this._owner.addChangedListener(this._ownerChangedListener);
                    }
                    this._owner = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Drive.prototype, "ownerChanged", {
                get: function () {
                    return this._ownerChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Drive.prototype, "quota", {
                /// <summary>
                /// There are no comments for Property quota in the schema.
                /// </summary>
                get: function () {
                    return this._quota;
                },
                set: function (value) {
                    if (this._quota) {
                        this._quota.removeChangedListener(this._quotaChangedListener);
                    }
                    if (value !== this._quota) {
                        this._quotaChanged = true;
                        this.changed = true;
                    }
                    if (this._quota) {
                        this._quota.addChangedListener(this._quotaChangedListener);
                    }
                    this._quota = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Drive.prototype, "quotaChanged", {
                get: function () {
                    return this._quotaChanged;
                },
                enumerable: true,
                configurable: true
            });

            Drive.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Drive.parseDrive(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Drive.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Drive.parseDrive = function (context, path, data) {
                if (!data)
                    return null;

                return new Drive(context, path, data);
            };

            Drive.parseDrives = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Drive.parseDrive(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Drive.prototype.getRequestBody = function () {
                return {
                    id: (this.idChanged && this.id) ? this.id : undefined,
                    owner: (this.ownerChanged && this.owner) ? this.owner.getRequestBody() : undefined,
                    quota: (this.quotaChanged && this.quota) ? this.quota.getRequestBody() : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Drive;
        })(Microsoft.CoreServices.Extensions.EntityBase);
        FileServices.Drive = Drive;

        /// <summary>
        /// There are no comments for Item in the schema.
        /// </summary>
        var ItemFetcher = (function (_super) {
            __extends(ItemFetcher, _super);
            function ItemFetcher(context, path) {
                _super.call(this, context, path);
            }
            return ItemFetcher;
        })(Microsoft.CoreServices.Extensions.RestShallowObjectFetcher);
        ItemFetcher.prototype.asFile = function () { return new FileFetcher(this.context, this.path); }; ItemFetcher.prototype.asFolder = function () { return new FolderFetcher(this.context, this.path); }; FileServices.ItemFetcher = ItemFetcher;

        /// <summary>
        /// There are no comments for Item in the schema.
        /// </summary>
        var Item = (function (_super) {
            __extends(Item, _super);
            function Item(context, path, data) {
                var _this = this;
                _super.call(this, context, path);
                this._odataType = 'Microsoft.FileServices.Item';
                this._createdByChanged = false;
                this._createdByChangedListener = (function (value) {
                    _this._createdByChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._eTagChanged = false;
                this._idChanged = false;
                this._lastModifiedByChanged = false;
                this._lastModifiedByChangedListener = (function (value) {
                    _this._lastModifiedByChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._nameChanged = false;
                this._parentReferenceChanged = false;
                this._parentReferenceChangedListener = (function (value) {
                    _this._parentReferenceChanged = true;
                    _this.changed = true;
                }).bind(this);
                this._sizeChanged = false;
                this._dateTimeCreatedChanged = false;
                this._dateTimeLastModifiedChanged = false;
                this._typeChanged = false;
                this._webUrlChanged = false;

                if (!data) {
                    return;
                }

                this._createdBy = IdentitySet.parseIdentitySet(data.createdBy);
                if (this._createdBy) {
                    this._createdBy.addChangedListener(this._createdByChangedListener);
                }
                this._eTag = data.eTag;
                this._id = data.id;
                this._lastModifiedBy = IdentitySet.parseIdentitySet(data.lastModifiedBy);
                if (this._lastModifiedBy) {
                    this._lastModifiedBy.addChangedListener(this._lastModifiedByChangedListener);
                }
                this._name = data.name;
                this._parentReference = ItemReference.parseItemReference(data.parentReference);
                if (this._parentReference) {
                    this._parentReference.addChangedListener(this._parentReferenceChangedListener);
                }
                this._size = data.size;
                this._dateTimeCreated = (data.dateTimeCreated !== null) ? new Date(data.dateTimeCreated) : null;
                this._dateTimeLastModified = (data.dateTimeLastModified !== null) ? new Date(data.dateTimeLastModified) : null;
                this._type = data.type;
                this._webUrl = data.webUrl;
            }
            Object.defineProperty(Item.prototype, "createdBy", {
                /// <summary>
                /// There are no comments for Property createdBy in the schema.
                /// </summary>
                get: function () {
                    return this._createdBy;
                },
                set: function (value) {
                    if (this._createdBy) {
                        this._createdBy.removeChangedListener(this._createdByChangedListener);
                    }
                    if (value !== this._createdBy) {
                        this._createdByChanged = true;
                        this.changed = true;
                    }
                    if (this._createdBy) {
                        this._createdBy.addChangedListener(this._createdByChangedListener);
                    }
                    this._createdBy = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "createdByChanged", {
                get: function () {
                    return this._createdByChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "eTag", {
                /// <summary>
                /// There are no comments for Property eTag in the schema.
                /// </summary>
                get: function () {
                    return this._eTag;
                },
                set: function (value) {
                    if (value !== this._eTag) {
                        this._eTagChanged = true;
                        this.changed = true;
                    }
                    this._eTag = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "eTagChanged", {
                get: function () {
                    return this._eTagChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "id", {
                /// <summary>
                /// There are no comments for Property id in the schema.
                /// </summary>
                get: function () {
                    return this._id;
                },
                set: function (value) {
                    if (value !== this._id) {
                        this._idChanged = true;
                        this.changed = true;
                    }
                    this._id = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "idChanged", {
                get: function () {
                    return this._idChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "lastModifiedBy", {
                /// <summary>
                /// There are no comments for Property lastModifiedBy in the schema.
                /// </summary>
                get: function () {
                    return this._lastModifiedBy;
                },
                set: function (value) {
                    if (this._lastModifiedBy) {
                        this._lastModifiedBy.removeChangedListener(this._lastModifiedByChangedListener);
                    }
                    if (value !== this._lastModifiedBy) {
                        this._lastModifiedByChanged = true;
                        this.changed = true;
                    }
                    if (this._lastModifiedBy) {
                        this._lastModifiedBy.addChangedListener(this._lastModifiedByChangedListener);
                    }
                    this._lastModifiedBy = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "lastModifiedByChanged", {
                get: function () {
                    return this._lastModifiedByChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "name", {
                /// <summary>
                /// There are no comments for Property name in the schema.
                /// </summary>
                get: function () {
                    return this._name;
                },
                set: function (value) {
                    if (value !== this._name) {
                        this._nameChanged = true;
                        this.changed = true;
                    }
                    this._name = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "nameChanged", {
                get: function () {
                    return this._nameChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "parentReference", {
                /// <summary>
                /// There are no comments for Property parentReference in the schema.
                /// </summary>
                get: function () {
                    return this._parentReference;
                },
                set: function (value) {
                    if (this._parentReference) {
                        this._parentReference.removeChangedListener(this._parentReferenceChangedListener);
                    }
                    if (value !== this._parentReference) {
                        this._parentReferenceChanged = true;
                        this.changed = true;
                    }
                    if (this._parentReference) {
                        this._parentReference.addChangedListener(this._parentReferenceChangedListener);
                    }
                    this._parentReference = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "parentReferenceChanged", {
                get: function () {
                    return this._parentReferenceChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "size", {
                /// <summary>
                /// There are no comments for Property size in the schema.
                /// </summary>
                get: function () {
                    return this._size;
                },
                set: function (value) {
                    if (value !== this._size) {
                        this._sizeChanged = true;
                        this.changed = true;
                    }
                    this._size = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "sizeChanged", {
                get: function () {
                    return this._sizeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "dateTimeCreated", {
                /// <summary>
                /// There are no comments for Property dateTimeCreated in the schema.
                /// </summary>
                get: function () {
                    return this._dateTimeCreated;
                },
                set: function (value) {
                    if (value !== this._dateTimeCreated) {
                        this._dateTimeCreatedChanged = true;
                        this.changed = true;
                    }
                    this._dateTimeCreated = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "dateTimeCreatedChanged", {
                get: function () {
                    return this._dateTimeCreatedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "dateTimeLastModified", {
                /// <summary>
                /// There are no comments for Property dateTimeLastModified in the schema.
                /// </summary>
                get: function () {
                    return this._dateTimeLastModified;
                },
                set: function (value) {
                    if (value !== this._dateTimeLastModified) {
                        this._dateTimeLastModifiedChanged = true;
                        this.changed = true;
                    }
                    this._dateTimeLastModified = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "dateTimeLastModifiedChanged", {
                get: function () {
                    return this._dateTimeLastModifiedChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "type", {
                /// <summary>
                /// There are no comments for Property type in the schema.
                /// </summary>
                get: function () {
                    return this._type;
                },
                set: function (value) {
                    if (value !== this._type) {
                        this._typeChanged = true;
                        this.changed = true;
                    }
                    this._type = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "typeChanged", {
                get: function () {
                    return this._typeChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Item.prototype, "webUrl", {
                /// <summary>
                /// There are no comments for Property webUrl in the schema.
                /// </summary>
                get: function () {
                    return this._webUrl;
                },
                set: function (value) {
                    if (value !== this._webUrl) {
                        this._webUrlChanged = true;
                        this.changed = true;
                    }
                    this._webUrl = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Item.prototype, "webUrlChanged", {
                get: function () {
                    return this._webUrlChanged;
                },
                enumerable: true,
                configurable: true
            });

            Item.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Item.parseItem(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Item.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Item.parseItem = function (context, path, data) {
                if (!data)
                    return null;

                if (data['@odata.type']) {
                    if (data['@odata.type'] === 'Microsoft.FileServices.File')
                        return new File(context, path, data);
                    if (data['@odata.type'] === 'Microsoft.FileServices.Folder')
                        return new Folder(context, path, data);
                }

                return new Item(context, path, data);
            };

            Item.parseItems = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Item.parseItem(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Item.prototype.getRequestBody = function () {
                return {
                    createdBy: (this.createdByChanged && this.createdBy) ? this.createdBy.getRequestBody() : undefined,
                    eTag: (this.eTagChanged && this.eTag) ? this.eTag : undefined,
                    id: (this.idChanged && this.id) ? this.id : undefined,
                    lastModifiedBy: (this.lastModifiedByChanged && this.lastModifiedBy) ? this.lastModifiedBy.getRequestBody() : undefined,
                    name: (this.nameChanged && this.name) ? this.name : undefined,
                    parentReference: (this.parentReferenceChanged && this.parentReference) ? this.parentReference.getRequestBody() : undefined,
                    size: (this.sizeChanged && this.size) ? this.size : undefined,
                    dateTimeCreated: (this.dateTimeCreatedChanged && this.dateTimeCreated) ? this.dateTimeCreated.toString() : undefined,
                    dateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    type: (this.typeChanged && this.type) ? this.type : undefined,
                    webUrl: (this.webUrlChanged && this.webUrl) ? this.webUrl : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Item;
        })(Microsoft.CoreServices.Extensions.EntityBase);
        FileServices.Item = Item;

        /// <summary>
        /// There are no comments for File in the schema.
        /// </summary>
        var FileFetcher = (function (_super) {
            __extends(FileFetcher, _super);
            function FileFetcher(context, path) {
                _super.call(this, context, path);
            }
            FileFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(File.parseFile(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            FileFetcher.prototype.content = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("content")); request.responseType = 'buffer';

                request.method = 'GET';

                this.context.request(request).then((function (data) {
                    deferred.resolve(data);
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            FileFetcher.prototype.copy = function (destFolderId, destFolderPath, newName) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("copy"));

                request.method = 'POST';
                request.data = JSON.stringify({ "destFolderId": destFolderId, "destFolderPath": destFolderPath, "newName": newName });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(File.parseFile(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            FileFetcher.prototype.uploadContent = function (contentStream) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("uploadContent"));

                request.method = 'POST';
                request.data = contentStream;

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };
            return FileFetcher;
        })(ItemFetcher);
        FileServices.FileFetcher = FileFetcher;

        /// <summary>
        /// There are no comments for File in the schema.
        /// </summary>
        var File = (function (_super) {
            __extends(File, _super);
            function File(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.FileServices.File';
                this._contentUrlChanged = false;

                if (!data) {
                    return;
                }

                this._contentUrl = data.contentUrl;
            }
            Object.defineProperty(File.prototype, "contentUrl", {
                /// <summary>
                /// There are no comments for Property contentUrl in the schema.
                /// </summary>
                get: function () {
                    return this._contentUrl;
                },
                set: function (value) {
                    if (value !== this._contentUrl) {
                        this._contentUrlChanged = true;
                        this.changed = true;
                    }
                    this._contentUrl = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(File.prototype, "contentUrlChanged", {
                get: function () {
                    return this._contentUrlChanged;
                },
                enumerable: true,
                configurable: true
            });

            File.prototype.content = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("content")); request.responseType = 'buffer';

                request.method = 'GET';

                this.context.request(request).then((function (data) {
                    deferred.resolve(data);
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            File.prototype.copy = function (destFolderId, destFolderPath, newName) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("copy"));

                request.method = 'POST';
                request.data = JSON.stringify({ "destFolderId": destFolderId, "destFolderPath": destFolderPath, "newName": newName });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(File.parseFile(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            File.prototype.uploadContent = function (contentStream) {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("uploadContent"));

                request.method = 'POST';
                request.data = contentStream;

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            File.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(File.parseFile(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            File.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            File.parseFile = function (context, path, data) {
                if (!data)
                    return null;

                return new File(context, path, data);
            };

            File.parseFiles = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(File.parseFile(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            File.prototype.getRequestBody = function () {
                return {
                    contentUrl: (this.contentUrlChanged && this.contentUrl) ? this.contentUrl : undefined,
                    createdBy: (this.createdByChanged && this.createdBy) ? this.createdBy.getRequestBody() : undefined,
                    eTag: (this.eTagChanged && this.eTag) ? this.eTag : undefined,
                    id: (this.idChanged && this.id) ? this.id : undefined,
                    lastModifiedBy: (this.lastModifiedByChanged && this.lastModifiedBy) ? this.lastModifiedBy.getRequestBody() : undefined,
                    name: (this.nameChanged && this.name) ? this.name : undefined,
                    parentReference: (this.parentReferenceChanged && this.parentReference) ? this.parentReference.getRequestBody() : undefined,
                    size: (this.sizeChanged && this.size) ? this.size : undefined,
                    dateTimeCreated: (this.dateTimeCreatedChanged && this.dateTimeCreated) ? this.dateTimeCreated.toString() : undefined,
                    dateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    type: (this.typeChanged && this.type) ? this.type : undefined,
                    webUrl: (this.webUrlChanged && this.webUrl) ? this.webUrl : undefined,
                    '@odata.type': this._odataType
                };
            };
            return File;
        })(Item);
        FileServices.File = File;

        /// <summary>
        /// There are no comments for Folder in the schema.
        /// </summary>
        var FolderFetcher = (function (_super) {
            __extends(FolderFetcher, _super);
            function FolderFetcher(context, path) {
                _super.call(this, context, path);
            }
            Object.defineProperty(FolderFetcher.prototype, "children", {
                get: function () {
                    if (this._children === undefined) {
                        this._children = new Microsoft.FileServices.Items(this.context, this.getPath('children'));
                    }
                    return this._children;
                },
                enumerable: true,
                configurable: true
            });

            FolderFetcher.prototype.fetch = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                this.context.readUrl(this.path).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            FolderFetcher.prototype.copy = function (destFolderId, destFolderPath, newName) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("copy"));

                request.method = 'POST';
                request.data = JSON.stringify({ "destFolderId": destFolderId, "destFolderPath": destFolderPath, "newName": newName });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return FolderFetcher;
        })(ItemFetcher);
        FileServices.FolderFetcher = FolderFetcher;

        /// <summary>
        /// There are no comments for Folder in the schema.
        /// </summary>
        var Folder = (function (_super) {
            __extends(Folder, _super);
            function Folder(context, path, data) {
                _super.call(this, context, path, data);
                this._odataType = 'Microsoft.FileServices.Folder';
                this._childCountChanged = false;

                if (!data) {
                    return;
                }

                this._childCount = data.childCount;
            }
            Object.defineProperty(Folder.prototype, "childCount", {
                /// <summary>
                /// There are no comments for Property childCount in the schema.
                /// </summary>
                get: function () {
                    return this._childCount;
                },
                set: function (value) {
                    if (value !== this._childCount) {
                        this._childCountChanged = true;
                        this.changed = true;
                    }
                    this._childCount = value;
                },
                enumerable: true,
                configurable: true
            });


            Object.defineProperty(Folder.prototype, "childCountChanged", {
                get: function () {
                    return this._childCountChanged;
                },
                enumerable: true,
                configurable: true
            });

            Object.defineProperty(Folder.prototype, "children", {
                get: function () {
                    if (this._children === undefined) {
                        this._children = new Microsoft.FileServices.Items(this.context, this.getPath('children'));
                    }
                    return this._children;
                },
                enumerable: true,
                configurable: true
            });

            Folder.prototype.copy = function (destFolderId, destFolderPath, newName) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("copy"));

                request.method = 'POST';
                request.data = JSON.stringify({ "destFolderId": destFolderId, "destFolderPath": destFolderPath, "newName": newName });

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };

            Folder.prototype.update = function () {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'PATCH';
                request.data = JSON.stringify(this.getRequestBody());

                this.context.request(request).then(function (data) {
                    var parsedData = JSON.parse(data);
                    deferred.resolve(Folder.parseFolder(_this.context, parsedData['@odata.id'], parsedData));
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Folder.prototype.delete = function () {
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.path);

                request.method = 'DELETE';

                this.context.request(request).then(function (data) {
                    deferred.resolve(null);
                }, deferred.reject.bind(deferred));

                return deferred;
            };

            Folder.parseFolder = function (context, path, data) {
                if (!data)
                    return null;

                return new Folder(context, path, data);
            };

            Folder.parseFolders = function (context, pathFn, data) {
                var results = [];

                if (data) {
                    for (var i = 0; i < data.length; ++i) {
                        results.push(Folder.parseFolder(context, pathFn(data[i]), data[i]));
                    }
                }

                return results;
            };

            Folder.prototype.getRequestBody = function () {
                return {
                    childCount: (this.childCountChanged && this.childCount) ? this.childCount : undefined,
                    createdBy: (this.createdByChanged && this.createdBy) ? this.createdBy.getRequestBody() : undefined,
                    eTag: (this.eTagChanged && this.eTag) ? this.eTag : undefined,
                    id: (this.idChanged && this.id) ? this.id : undefined,
                    lastModifiedBy: (this.lastModifiedByChanged && this.lastModifiedBy) ? this.lastModifiedBy.getRequestBody() : undefined,
                    name: (this.nameChanged && this.name) ? this.name : undefined,
                    parentReference: (this.parentReferenceChanged && this.parentReference) ? this.parentReference.getRequestBody() : undefined,
                    size: (this.sizeChanged && this.size) ? this.size : undefined,
                    dateTimeCreated: (this.dateTimeCreatedChanged && this.dateTimeCreated) ? this.dateTimeCreated.toString() : undefined,
                    dateTimeLastModified: (this.dateTimeLastModifiedChanged && this.dateTimeLastModified) ? this.dateTimeLastModified.toString() : undefined,
                    type: (this.typeChanged && this.type) ? this.type : undefined,
                    webUrl: (this.webUrlChanged && this.webUrl) ? this.webUrl : undefined,
                    '@odata.type': this._odataType
                };
            };
            return Folder;
        })(Item);
        FileServices.Folder = Folder;
        var Items = (function (_super) {
            __extends(Items, _super);
            function Items(context, path, entity) {
                _super.call(this, context, path, entity);

                this._parseCollectionFn = function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return Item.parseItems(context, pathFn, data.value);
                };
            }
            Items.prototype.getItem = function (id) {
                var path = this.path + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name: "id", type: "Edm.String", value: id }]);
                var fetcher = new ItemFetcher(this.context, path);
                return fetcher;
            };

            Items.prototype.getItems = function () {
                return new Microsoft.CoreServices.Extensions.CollectionQuery(this.context, this.path, this._parseCollectionFn);
            };

            Items.prototype.addItem = function (item) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred();

                if (this.entity == null) {
                    var request = new Microsoft.CoreServices.Extensions.Request(this.path);

                    request.method = 'POST';
                    request.data = JSON.stringify(item.getRequestBody());

                    this.context.request(request).then((function (data) {
                        var parsedData = JSON.parse(data);
                        deferred.resolve(Item.parseItem(_this.context, parsedData['@odata.id'], parsedData));
                    }).bind(this), deferred.reject.bind(deferred));
                } else {
                    //UNDONE this.context.AddLink(_entity, _path, item);
                }

                return deferred;
            };
            Items.prototype.asFiles = function () {
                var parseCollectionFn = (function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return File.parseFiles(context, pathFn, data.value);
                }).bind(this);
                return new Microsoft.CoreServices.Extensions.CollectionQuery(this.context, this.path + '/$/Microsoft.FileServices.File()', parseCollectionFn);
            };
            Items.prototype.asFolders = function () {
                var parseCollectionFn = (function (context, data) {
                    var pathFn = function (data) {
                        return data['@odata.id'];
                    };
                    return Folder.parseFolders(context, pathFn, data.value);
                }).bind(this);
                return new Microsoft.CoreServices.Extensions.CollectionQuery(this.context, this.path + '/$/Microsoft.FileServices.Folder()', parseCollectionFn);
            };
            Items.prototype.getByPath = function (path) {
                var _this = this;
                var deferred = new Microsoft.Utility.Deferred(), request = new Microsoft.CoreServices.Extensions.Request(this.getPath("getByPath"));

                request.method = 'GET';

                this.context.request(request).then((function (data) {
                    var parsedData = JSON.parse(data);

                    //var path = this.context.serviceRootUri + '/items' + Microsoft.Utility.EncodingHelpers.getKeyExpression([{ name : "id", type : "Edm.String", value : parsedData.d.id }]);
                    var path = data.d['__metadata']['id'];
                    deferred.resolve(Item.parseItem(_this.context, path, parsedData.d));
                }).bind(this), deferred.reject.bind(deferred));

                return deferred;
            };
            return Items;
        })(Microsoft.CoreServices.Extensions.QueryableSet);
        FileServices.Items = Items;
    })(Microsoft.FileServices || (Microsoft.FileServices = {}));
    var FileServices = Microsoft.FileServices;
})(Microsoft || (Microsoft = {}));
//# sourceMappingURL=sharepoint.js.map
module.exports = { Microsoft: Microsoft, O365Auth: O365Auth, O365Discovery: O365Discovery };
