define("dd7ec4fd-97aa-44c5-b6ad-87535862e0bf_2.0.0", ["announcementsStrings","@microsoft/sp-application-base","@microsoft/office-ui-fabric-react-bundle","@microsoft/sp-core-library","react","react-dom","@microsoft/decorators"], function(__WEBPACK_EXTERNAL_MODULE__n6P__, __WEBPACK_EXTERNAL_MODULE_GPet__, __WEBPACK_EXTERNAL_MODULE_KL1q__, __WEBPACK_EXTERNAL_MODULE_UWqr__, __WEBPACK_EXTERNAL_MODULE_cDcd__, __WEBPACK_EXTERNAL_MODULE_faye__, __WEBPACK_EXTERNAL_MODULE_wxtz__) { return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "wdEC");
/******/ })
/************************************************************************/
/******/ ({

/***/ "+QQQ":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/profiles/types.js ***!
  \************************************************/
/*! exports provided: _Profiles, Profiles, UrlZone */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Profiles", function() { return _Profiles; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Profiles", function() { return Profiles; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "UrlZone", function() { return UrlZone; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../operations */ "UK2s");







var _Profiles = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Profiles, _super);
    /**
     * Creates a new instance of the UserProfileQuery class
     *
     * @param baseUrl The url or SharePointQueryable which forms the parent of this user profile query
     */
    function _Profiles(baseUrl, path) {
        if (path === void 0) { path = "_api/sp.userprofiles.peoplemanager"; }
        var _this = _super.call(this, baseUrl, path) || this;
        _this.clientPeoplePickerQuery = (new ClientPeoplePickerQuery(baseUrl)).configureFrom(_this);
        _this.profileLoader = (new ProfileLoader(baseUrl)).configureFrom(_this);
        return _this;
    }
    Object.defineProperty(_Profiles.prototype, "editProfileLink", {
        /**
         * The url of the edit profile page for the current user
         */
        get: function () {
            return this.clone(Profiles, "EditProfileLink").get();
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Profiles.prototype, "isMyPeopleListPublic", {
        /**
         * A boolean value that indicates whether the current user's "People I'm Following" list is public
         */
        get: function () {
            return this.clone(Profiles, "IsMyPeopleListPublic").get();
        },
        enumerable: true,
        configurable: true
    });
    /**
     * A boolean value that indicates whether the current user is being followed by the specified user
     *
     * @param loginName The account name of the user
     */
    _Profiles.prototype.amIFollowedBy = function (loginName) {
        var q = this.clone(Profiles, "amifollowedby(@v)");
        q.query.set("@v", "'" + encodeURIComponent(loginName) + "'");
        return q.get();
    };
    /**
     * A boolean value that indicates whether the current user is following the specified user
     *
     * @param loginName The account name of the user
     */
    _Profiles.prototype.amIFollowing = function (loginName) {
        var q = this.clone(Profiles, "amifollowing(@v)");
        q.query.set("@v", "'" + encodeURIComponent(loginName) + "'");
        return q.get();
    };
    /**
     * Gets tags that the current user is following
     *
     * @param maxCount The maximum number of tags to retrieve (default is 20)
     */
    _Profiles.prototype.getFollowedTags = function (maxCount) {
        if (maxCount === void 0) { maxCount = 20; }
        return this.clone(Profiles, "getfollowedtags(" + maxCount + ")").get();
    };
    /**
     * Gets the people who are following the specified user
     *
     * @param loginName The account name of the user
     */
    _Profiles.prototype.getFollowersFor = function (loginName) {
        var q = this.clone(Profiles, "getfollowersfor(@v)");
        q.query.set("@v", "'" + encodeURIComponent(loginName) + "'");
        return q.get();
    };
    Object.defineProperty(_Profiles.prototype, "myFollowers", {
        /**
         * Gets the people who are following the current user
         *
         */
        get: function () {
            return Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryableCollection"])(this, "getmyfollowers");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Profiles.prototype, "myProperties", {
        /**
         * Gets user properties for the current user
         *
         */
        get: function () {
            return new _Profiles(this, "getmyproperties");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Gets the people who the specified user is following
     *
     * @param loginName The account name of the user.
     */
    _Profiles.prototype.getPeopleFollowedBy = function (loginName) {
        var q = this.clone(Profiles, "getpeoplefollowedby(@v)");
        q.query.set("@v", "'" + encodeURIComponent(loginName) + "'");
        return q.get();
    };
    /**
     * Gets user properties for the specified user.
     *
     * @param loginName The account name of the user.
     */
    _Profiles.prototype.getPropertiesFor = function (loginName) {
        var q = this.clone(Profiles, "getpropertiesfor(@v)");
        q.query.set("@v", "'" + encodeURIComponent(loginName) + "'");
        return q.get();
    };
    Object.defineProperty(_Profiles.prototype, "trendingTags", {
        /**
         * Gets the 20 most popular hash tags over the past week, sorted so that the most popular tag appears first
         *
         */
        get: function () {
            var q = this.clone(Profiles, null);
            q.concat(".gettrendingtags");
            return q.get();
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Gets the specified user profile property for the specified user
     *
     * @param loginName The account name of the user
     * @param propertyName The case-sensitive name of the property to get
     */
    _Profiles.prototype.getUserProfilePropertyFor = function (loginName, propertyName) {
        var q = this.clone(Profiles, "getuserprofilepropertyfor(accountname=@v, propertyname='" + propertyName + "')");
        q.query.set("@v", "'" + encodeURIComponent(loginName) + "'");
        return q.get();
    };
    /**
     * Removes the specified user from the user's list of suggested people to follow
     *
     * @param loginName The account name of the user
     */
    _Profiles.prototype.hideSuggestion = function (loginName) {
        var q = this.clone(Profiles, "hidesuggestion(@v)");
        q.query.set("@v", "'" + encodeURIComponent(loginName) + "'");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(q);
    };
    /**
     * A boolean values that indicates whether the first user is following the second user
     *
     * @param follower The account name of the user who might be following the followee
     * @param followee The account name of the user who might be followed by the follower
     */
    _Profiles.prototype.isFollowing = function (follower, followee) {
        var q = this.clone(Profiles, null);
        q.concat(".isfollowing(possiblefolloweraccountname=@v, possiblefolloweeaccountname=@y)");
        q.query.set("@v", "'" + encodeURIComponent(follower) + "'");
        q.query.set("@y", "'" + encodeURIComponent(followee) + "'");
        return q.get();
    };
    /**
     * Uploads and sets the user profile picture (Users can upload a picture to their own profile only). Not supported for batching.
     *
     * @param profilePicSource Blob data representing the user's picture in BMP, JPEG, or PNG format of up to 4.76MB
     */
    _Profiles.prototype.setMyProfilePic = function (profilePicSource) {
        var buffer = null;
        var reader = new FileReader();
        reader.onload = function (e) { return buffer = e.target.result; };
        reader.readAsArrayBuffer(profilePicSource);
        var request = new _Profiles(this, "setmyprofilepicture");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(request, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])(String.fromCharCode.apply(null, new Uint16Array(buffer))));
    };
    /**
     * Sets single value User Profile property
     *
     * @param accountName The account name of the user
     * @param propertyName Property name
     * @param propertyValue Property value
     */
    _Profiles.prototype.setSingleValueProfileProperty = function (accountName, propertyName, propertyValue) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Profiles, "SetSingleValueProfileProperty"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({
            accountName: accountName,
            propertyName: propertyName,
            propertyValue: propertyValue,
        }));
    };
    /**
     * Sets multi valued User Profile property
     *
     * @param accountName The account name of the user
     * @param propertyName Property name
     * @param propertyValues Property values
     */
    _Profiles.prototype.setMultiValuedProfileProperty = function (accountName, propertyName, propertyValues) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Profiles, "SetMultiValuedProfileProperty"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({
            accountName: accountName,
            propertyName: propertyName,
            propertyValues: propertyValues,
        }));
    };
    /**
     * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only)
     *
     * @param emails The email addresses of the users to provision sites for
     */
    _Profiles.prototype.createPersonalSiteEnqueueBulk = function () {
        var emails = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            emails[_i] = arguments[_i];
        }
        return this.profileLoader.createPersonalSiteEnqueueBulk(emails);
    };
    Object.defineProperty(_Profiles.prototype, "ownerUserProfile", {
        /**
         * Gets the user profile of the site owner
         *
         */
        get: function () {
            return this.profileLoader.ownerUserProfile;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Profiles.prototype, "userProfile", {
        /**
         * Gets the user profile for the current user
         */
        get: function () {
            return this.profileLoader.userProfile;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files
     *
     * @param interactiveRequest true if interactively (web) initiated request, or false (default) if non-interactively (client) initiated request
     */
    _Profiles.prototype.createPersonalSite = function (interactiveRequest) {
        if (interactiveRequest === void 0) { interactiveRequest = false; }
        return this.profileLoader.createPersonalSite(interactiveRequest);
    };
    /**
     * Sets the privacy settings for this profile
     *
     * @param share true to make all social data public; false to make all social data private
     */
    _Profiles.prototype.shareAllSocialData = function (share) {
        return this.profileLoader.shareAllSocialData(share);
    };
    /**
     * Resolves user or group using specified query parameters
     *
     * @param queryParams The query parameters used to perform resolve
     */
    _Profiles.prototype.clientPeoplePickerResolveUser = function (queryParams) {
        return this.clientPeoplePickerQuery.clientPeoplePickerResolveUser(queryParams);
    };
    /**
     * Searches for users or groups using specified query parameters
     *
     * @param queryParams The query parameters used to perform search
     */
    _Profiles.prototype.clientPeoplePickerSearchUser = function (queryParams) {
        return this.clientPeoplePickerQuery.clientPeoplePickerSearchUser(queryParams);
    };
    return _Profiles;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Profiles = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Profiles);
var ProfileLoader = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(ProfileLoader, _super);
    function ProfileLoader() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Provisions one or more users' personal sites. (My Site administrator on SharePoint Online only) Doesn't support batching
     *
     * @param emails The email addresses of the users to provision sites for
     */
    ProfileLoader.prototype.createPersonalSiteEnqueueBulk = function (emails) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(ProfileLoaderFactory, "createpersonalsiteenqueuebulk", false), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({ "emailIDs": emails }));
    };
    Object.defineProperty(ProfileLoader.prototype, "ownerUserProfile", {
        /**
         * Gets the user profile of the site owner.
         *
         */
        get: function () {
            var q = this.getParent(ProfileLoaderFactory, this.parentUrl, "_api/sp.userprofiles.profileloader.getowneruserprofile");
            if (this.hasBatch) {
                q = q.inBatch(this.batch);
            }
            return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(q);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ProfileLoader.prototype, "userProfile", {
        /**
         * Gets the user profile of the current user.
         *
         */
        get: function () {
            return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(ProfileLoaderFactory, "getuserprofile"));
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Enqueues creating a personal site for this user, which can be used to share documents, web pages, and other files.
     *
     * @param interactiveRequest true if interactively (web) initiated request, or false (default) if non-interactively (client) initiated request
     */
    ProfileLoader.prototype.createPersonalSite = function (interactiveRequest) {
        if (interactiveRequest === void 0) { interactiveRequest = false; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(ProfileLoaderFactory, "getuserprofile/createpersonalsiteenque(" + interactiveRequest + ")"));
    };
    /**
     * Sets the privacy settings for this profile
     *
     * @param share true to make all social data public; false to make all social data private.
     */
    ProfileLoader.prototype.shareAllSocialData = function (share) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(ProfileLoaderFactory, "getuserprofile/shareallsocialdata(" + share + ")"));
    };
    ProfileLoader = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_5__["defaultPath"])("_api/sp.userprofiles.profileloader.getprofileloader")
    ], ProfileLoader);
    return ProfileLoader;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryable"]));
var ProfileLoaderFactory = function (baseUrl, path) {
    return new ProfileLoader(baseUrl, path);
};
var ClientPeoplePickerQuery = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(ClientPeoplePickerQuery, _super);
    function ClientPeoplePickerQuery() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Resolves user or group using specified query parameters
     *
     * @param queryParams The query parameters used to perform resolve
     */
    ClientPeoplePickerQuery.prototype.clientPeoplePickerResolveUser = function (queryParams) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var q, res;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        q = this.clone(ClientPeoplePickerFactory, null);
                        q.concat(".clientpeoplepickerresolveuser");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(q, this.getBodyFrom(queryParams))];
                    case 1:
                        res = _a.sent();
                        return [2 /*return*/, JSON.parse(typeof res === "object" ? res.ClientPeoplePickerResolveUser : res)];
                }
            });
        });
    };
    /**
     * Searches for users or groups using specified query parameters
     *
     * @param queryParams The query parameters used to perform search
     */
    ClientPeoplePickerQuery.prototype.clientPeoplePickerSearchUser = function (queryParams) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var q, res;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        q = this.clone(ClientPeoplePickerFactory, null);
                        q.concat(".clientpeoplepickersearchuser");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(q, this.getBodyFrom(queryParams))];
                    case 1:
                        res = _a.sent();
                        return [2 /*return*/, JSON.parse(typeof res === "object" ? res.ClientPeoplePickerSearchUser : res)];
                }
            });
        });
    };
    /**
     * Creates ClientPeoplePickerQueryParameters request body
     *
     * @param queryParams The query parameters to create request body
     */
    ClientPeoplePickerQuery.prototype.getBodyFrom = function (queryParams) {
        return Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({ "queryParams": Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])("SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters"), queryParams) });
    };
    ClientPeoplePickerQuery = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_5__["defaultPath"])("_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface")
    ], ClientPeoplePickerQuery);
    return ClientPeoplePickerQuery;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryable"]));
var ClientPeoplePickerFactory = function (baseUrl, path) {
    return new ClientPeoplePickerQuery(baseUrl, path);
};
/**
 * Specifies the originating zone of a request received.
 */
var UrlZone;
(function (UrlZone) {
    /**
     * Specifies the default zone used for requests unless another zone is specified.
     */
    UrlZone[UrlZone["DefaultZone"] = 0] = "DefaultZone";
    /**
     * Specifies an intranet zone.
     */
    UrlZone[UrlZone["Intranet"] = 1] = "Intranet";
    /**
     * Specifies an Internet zone.
     */
    UrlZone[UrlZone["Internet"] = 2] = "Internet";
    /**
     * Specifies a custom zone.
     */
    UrlZone[UrlZone["Custom"] = 3] = "Custom";
    /**
     * Specifies an extranet zone.
     */
    UrlZone[UrlZone["Extranet"] = 4] = "Extranet";
})(UrlZone || (UrlZone = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "+n6P":
/*!***************************************!*\
  !*** external "announcementsStrings" ***!
  \***************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE__n6P__;

/***/ }),

/***/ "01TU":
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/odata/invokable-extensions.js ***!
  \*********************************************************/
/*! exports provided: extendGlobal, extendObj, extendFactory, clearGlobalExtensions, disableExtensions, enableExtensions, applyFactoryExtensions, extensionOrDefault */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "extendGlobal", function() { return extendGlobal; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "extendObj", function() { return extendObj; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "extendFactory", function() { return extendFactory; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "clearGlobalExtensions", function() { return clearGlobalExtensions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "disableExtensions", function() { return disableExtensions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "enableExtensions", function() { return enableExtensions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "applyFactoryExtensions", function() { return applyFactoryExtensions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "extensionOrDefault", function() { return extensionOrDefault; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "t3D1");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");


var _enableExtensions = false;
var globaExtensions = [];
var ObjExtensionsSym = Symbol("__extensions");
/**
 * Creates global extensions across all invokable objects
 *
 * @param e The global extensions to apply
 */
var extendGlobal = function (e) {
    _enableExtensions = true;
    extendCol(globaExtensions, e);
};
/**
 * Applies the supplied extensions to a single instance
 *
 * @param target Object to which extensions are applied
 * @param extensions Extensions to apply
 */
var extendObj = function (target, extensions) {
    _enableExtensions = true;
    if (!Reflect.has(target, ObjExtensionsSym)) {
        Reflect.set(target, ObjExtensionsSym, []);
    }
    extendCol(Reflect.get(target, ObjExtensionsSym), extensions);
    return target;
};
/**
 * Allows applying extensions to all instances created from the supplied factory
 *
 * @param factory The Invokable Factory method to extend
 * @param extensions Extensions to apply
 */
var extendFactory = function (factory, extensions) {
    _enableExtensions = true;
    if (factory.__proto__[ObjExtensionsSym] === undefined) {
        factory.__proto__[ObjExtensionsSym] = [];
    }
    extendCol(factory.__proto__[ObjExtensionsSym], extensions);
};
function extendCol(a, e) {
    if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isArray"])(e)) {
        // @ts-ignore
        a.push.apply(a, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])(e));
    }
    else {
        // @ts-ignore
        a.push(e);
    }
}
/**
 * Clears all global extensions
 */
var clearGlobalExtensions = function () {
    globaExtensions.length = 0;
};
/**
 * Disables all extensions
 */
var disableExtensions = function () {
    _enableExtensions = false;
};
/**
 * Enables all extensions
 */
var enableExtensions = function () {
    _enableExtensions = true;
};
/**
 * Applies a set of extension previously applied to a factory using extendFactory to an object created from that factory
 *
 * @param factory
 * @param args
 */
var applyFactoryExtensions = function (factory, args) {
    var o = factory(args);
    if (factory.__proto__[ObjExtensionsSym]) {
        o = extendObj(o, factory.__proto__[ObjExtensionsSym]);
    }
    return o;
};
function extensionOrDefault(op, or, target) {
    var rest = [];
    for (var _i = 3; _i < arguments.length; _i++) {
        rest[_i - 3] = arguments[_i];
    }
    if (_enableExtensions) {
        var extensions = [];
        // we need to first invoke extensions tied to only this object
        if (Reflect.has(target, ObjExtensionsSym)) {
            extensions.push.apply(extensions, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])(Reflect.get(target, ObjExtensionsSym)));
        }
        // second we need to process any global extensions
        extensions.push.apply(extensions, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])(globaExtensions));
        for (var i = 0; i < extensions.length; i++) {
            var extension = extensions[i];
            var result = undefined;
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isFunc"])(extension)) {
                // this extension is a function which we call
                result = extension.apply(void 0, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])([op, target], rest));
            }
            else if (op === "get" && Reflect.has(extension, rest[0])) {
                // this extension is a named extension meaning we are overriding a specific method/property
                result = Reflect.get(extension, rest[0], target);
            }
            else if (Reflect.has(extension, op)) {
                // this extension is a ProxyHandler that has a handler defined for {op} so we pass control and see if we get a result
                result = Reflect.get(extension, op).apply(void 0, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])([target], rest));
            }
            if (typeof result !== "undefined") {
                // if a extension returned a result, we return that
                // this means that this extension overrides any other extensions and no more are executed
                // first extension in the list to return "wins"
                return result;
            }
        }
    }
    return or.apply(void 0, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])([target], rest));
}
//# sourceMappingURL=invokable-extensions.js.map

/***/ }),

/***/ "08ZC":
/*!*************************************************!*\
  !*** ./node_modules/@pnp/common/safe-global.js ***!
  \*************************************************/
/*! exports provided: safeGlobal */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* WEBPACK VAR INJECTION */(function(global) {/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "safeGlobal", function() { return safeGlobal; });
// export either window or global
var safeGlobal = typeof global === "undefined" ? window : global;
//# sourceMappingURL=safe-global.js.map
/* WEBPACK VAR INJECTION */}.call(this, __webpack_require__(/*! ./../../webpack/buildin/global.js */ "yLpj")))

/***/ }),

/***/ "0FKx":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/appcatalog/index.js ***!
  \**************************************************/
/*! exports provided: App, AppCatalog */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./web */ "amhO");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./types */ "ArVI");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "App", function() { return _types__WEBPACK_IMPORTED_MODULE_5__["App"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "AppCatalog", function() { return _types__WEBPACK_IMPORTED_MODULE_5__["AppCatalog"]; });







_rest__WEBPACK_IMPORTED_MODULE_1__["SPRest"].prototype.getTenantAppCatalogWeb = function () {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var data;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_4__["SharePointQueryable"])("/", "_api/SP_TenantSettings_Current")()];
                case 1:
                    data = _a.sent();
                    return [2 /*return*/, Object(_webs_types__WEBPACK_IMPORTED_MODULE_2__["Web"])(data.CorporateCatalogUrl)];
            }
        });
    });
};
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "0MvI":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/index.js ***!
  \***********************************************/
/*! exports provided: Folder, Folders */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _item__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./item */ "LZ64");
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./list */ "vttn");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./web */ "ymW/");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "4tXH");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Folder", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["Folder"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Folders", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["Folders"]; });





//# sourceMappingURL=index.js.map

/***/ }),

/***/ "0YVK":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/files/index.js ***!
  \*********************************************/
/*! exports provided: File, Files, CheckinType, MoveOperations, TemplateFileType, Version, Versions */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _folder__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./folder */ "JFDc");
/* harmony import */ var _item__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./item */ "pGs8");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./web */ "SPHc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "qfcF");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "File", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["File"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Files", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["Files"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CheckinType", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["CheckinType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "MoveOperations", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["MoveOperations"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "TemplateFileType", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["TemplateFileType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Version", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["Version"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Versions", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["Versions"]; });





//# sourceMappingURL=index.js.map

/***/ }),

/***/ "0p/9":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/splibconfig.js ***!
  \*********************************************/
/*! exports provided: emptyGuid, setup, SPRuntimeConfigImpl, SPRuntimeConfig */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "emptyGuid", function() { return emptyGuid; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setup", function() { return setup; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SPRuntimeConfigImpl", function() { return SPRuntimeConfigImpl; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SPRuntimeConfig", function() { return SPRuntimeConfig; });
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/common */ "zqEK");

var emptyGuid = "00000000-0000-0000-0000-000000000000";
function setup(config) {
    _pnp_common__WEBPACK_IMPORTED_MODULE_0__["RuntimeConfig"].assign(config);
}
var SPRuntimeConfigImpl = /** @class */ (function () {
    function SPRuntimeConfigImpl() {
    }
    Object.defineProperty(SPRuntimeConfigImpl.prototype, "headers", {
        get: function () {
            var spPart = _pnp_common__WEBPACK_IMPORTED_MODULE_0__["RuntimeConfig"].get("sp");
            if (spPart !== undefined && spPart.headers !== undefined) {
                return spPart.headers;
            }
            return {};
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRuntimeConfigImpl.prototype, "baseUrl", {
        get: function () {
            var spPart = _pnp_common__WEBPACK_IMPORTED_MODULE_0__["RuntimeConfig"].get("sp");
            if (spPart !== undefined && spPart.baseUrl !== undefined) {
                return spPart.baseUrl;
            }
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["objectDefinedNotNull"])(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["RuntimeConfig"].spfxContext)) {
                return _pnp_common__WEBPACK_IMPORTED_MODULE_0__["RuntimeConfig"].spfxContext.pageContext.web.absoluteUrl;
            }
            return null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRuntimeConfigImpl.prototype, "fetchClientFactory", {
        get: function () {
            var spPart = _pnp_common__WEBPACK_IMPORTED_MODULE_0__["RuntimeConfig"].get("sp");
            if (spPart !== undefined && spPart.fetchClientFactory !== undefined) {
                return spPart.fetchClientFactory;
            }
            else {
                return function () { return new _pnp_common__WEBPACK_IMPORTED_MODULE_0__["FetchClient"](); };
            }
        },
        enumerable: true,
        configurable: true
    });
    return SPRuntimeConfigImpl;
}());

var SPRuntimeConfig = new SPRuntimeConfigImpl();
//# sourceMappingURL=splibconfig.js.map

/***/ }),

/***/ "0rHm":
/*!*************************************************!*\
  !*** ./node_modules/@pnp/sp/site-groups/web.js ***!
  \*************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "UuUm");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");
/* harmony import */ var _security_web__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../security/web */ "4G66");







Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "siteGroups", _types__WEBPACK_IMPORTED_MODULE_3__["SiteGroups"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "associatedOwnerGroup", _types__WEBPACK_IMPORTED_MODULE_3__["SiteGroups"], "associatedownergroup");
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "associatedMemberGroup", _types__WEBPACK_IMPORTED_MODULE_3__["SiteGroups"], "associatedmembergroup");
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "associatedVisitorGroup", _types__WEBPACK_IMPORTED_MODULE_3__["SiteGroups"], "associatedvisitorgroup");
_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"].prototype.createDefaultAssociatedGroups = function (groupNameSeed, siteOwner, copyRoleAssignments, clearSubscopes, siteOwner2) {
    if (copyRoleAssignments === void 0) { copyRoleAssignments = false; }
    if (clearSubscopes === void 0) { clearSubscopes = true; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var q;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, this.breakRoleInheritance(copyRoleAssignments, clearSubscopes)];
                case 1:
                    _a.sent();
                    q = this.clone(_webs_types__WEBPACK_IMPORTED_MODULE_2__["Web"], "createDefaultAssociatedGroups(userLogin=@u,userLogin2=@v,groupNameSeed=@s)");
                    q.query.set("@u", "'" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_5__["escapeQueryStrValue"])(siteOwner || "") + "'");
                    q.query.set("@v", "'" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_5__["escapeQueryStrValue"])(siteOwner2 || "") + "'");
                    q.query.set("@s", "'" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_5__["escapeQueryStrValue"])(groupNameSeed || "") + "'");
                    return [2 /*return*/, Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(q)];
            }
        });
    });
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "1dmT":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/security/types.js ***!
  \************************************************/
/*! exports provided: _RoleAssignments, RoleAssignments, _RoleAssignment, RoleAssignment, _RoleDefinitions, RoleDefinitions, _RoleDefinition, RoleDefinition, PermissionKind */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_RoleAssignments", function() { return _RoleAssignments; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RoleAssignments", function() { return RoleAssignments; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_RoleAssignment", function() { return _RoleAssignment; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RoleAssignment", function() { return RoleAssignment; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_RoleDefinitions", function() { return _RoleDefinitions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RoleDefinitions", function() { return RoleDefinitions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_RoleDefinition", function() { return _RoleDefinition; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RoleDefinition", function() { return RoleDefinition; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PermissionKind", function() { return PermissionKind; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _site_groups_types__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../site-groups/types */ "UuUm");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../telemetry */ "h3gW");









/**
 * Describes a set of role assignments for the current scope
 *
 */
var _RoleAssignments = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_RoleAssignments, _super);
    function _RoleAssignments() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets the role assignment associated with the specified principal id from the collection.
     *
     * @param id The id of the role assignment
     */
    _RoleAssignments.prototype.getById = function (id) {
        return RoleAssignment(this).concat("(" + id + ")");
    };
    /**
     * Adds a new role assignment with the specified principal and role definitions to the collection
     *
     * @param principalId The id of the user or group to assign permissions to
     * @param roleDefId The id of the role definition that defines the permissions to assign
     *
     */
    _RoleAssignments.prototype.add = function (principalId, roleDefId) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(RoleAssignments, "addroleassignment(principalid=" + principalId + ", roledefid=" + roleDefId + ")"))];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Removes the role assignment with the specified principal and role definition from the collection
     *
     * @param principalId The id of the user or group in the role assignment
     * @param roleDefId The id of the role definition in the role assignment
     *
     */
    _RoleAssignments.prototype.remove = function (principalId, roleDefId) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(RoleAssignments, "removeroleassignment(principalid=" + principalId + ", roledefid=" + roleDefId + ")"))];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    _RoleAssignments = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("roleassignments")
    ], _RoleAssignments);
    return _RoleAssignments;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableCollection"]));

var RoleAssignments = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_RoleAssignments);
/**
 * Describes a role assignment
 *
 */
var _RoleAssignment = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_RoleAssignment, _super);
    function _RoleAssignment() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["deleteable"])("ra");
        return _this;
    }
    Object.defineProperty(_RoleAssignment.prototype, "groups", {
        /**
         * Gets the groups that directly belong to the access control list (ACL) for this securable object
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(Object(_site_groups_types__WEBPACK_IMPORTED_MODULE_4__["SiteGroups"])(this, "groups"), "ra.groups");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_RoleAssignment.prototype, "bindings", {
        /**
         * Gets the role definition bindings for this role assignment
         *
         */
        get: function () {
            return Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableCollection"])(this, "roledefinitionbindings");
        },
        enumerable: true,
        configurable: true
    });
    return _RoleAssignment;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableInstance"]));

var RoleAssignment = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_RoleAssignment);
/**
 * Describes a collection of role definitions
 *
 */
var _RoleDefinitions = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_RoleDefinitions, _super);
    function _RoleDefinitions() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets the role definition with the specified id from the collection
     *
     * @param id The id of the role definition
     *
     */
    _RoleDefinitions.prototype.getById = function (id) {
        return RoleDefinition(this, "getById(" + id + ")");
    };
    /**
     * Gets the role definition with the specified name
     *
     * @param name The name of the role definition
     *
     */
    _RoleDefinitions.prototype.getByName = function (name) {
        return RoleDefinition(this, "getbyname('" + name + "')");
    };
    /**
     * Gets the role definition with the specified role type
     *
     * @param roleTypeKind The roletypekind of the role definition (None=0, Guest=1, Reader=2, Contributor=3, WebDesigner=4, Administrator=5, Editor=6, System=7)
     *
     */
    _RoleDefinitions.prototype.getByType = function (roleTypeKind) {
        return RoleDefinition(this, "getbytype(" + roleTypeKind + ")");
    };
    /**
     * Creates a role definition
     *
     * @param name The new role definition's name
     * @param description The new role definition's description
     * @param order The order in which the role definition appears
     * @param basePermissions The permissions mask for this role definition
     *
     */
    _RoleDefinitions.prototype.add = function (name, description, order, basePermissions) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({
                            BasePermissions: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.BasePermissions"), basePermissions),
                            Description: description,
                            Name: name,
                            Order: order,
                            __metadata: { "type": "SP.RoleDefinition" },
                        });
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this, postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                definition: this.getById(data.Id),
                            }];
                }
            });
        });
    };
    _RoleDefinitions = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("roledefinitions")
    ], _RoleDefinitions);
    return _RoleDefinitions;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableCollection"]));

var RoleDefinitions = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_RoleDefinitions);
/**
 * Describes a role definition
 *
 */
var _RoleDefinition = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_RoleDefinition, _super);
    function _RoleDefinition() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["deleteable"])("rd");
        return _this;
        /* tslint:enable */
    }
    /**
     * Updates this role definition with the supplied properties
     *
     * @param properties A plain object hash of values to update for the role definition
     */
    /* tslint:disable no-string-literal */
    _RoleDefinition.prototype.update = function (properties) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var s, postBody, data, definition, parent_1;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        s = ["BasePermissions"];
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(properties, s[0]) !== undefined) {
                            properties[s[0]] = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP." + s[0]), properties[s[0]]);
                        }
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.RoleDefinition"), properties), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "X-HTTP-Method": "MERGE" }));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this, postBody)];
                    case 1:
                        data = _a.sent();
                        definition = this;
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(properties, "Name")) {
                            parent_1 = this.getParent(RoleDefinitions, this.parentUrl, "");
                            definition = parent_1.getByName(properties["Name"]);
                        }
                        return [2 /*return*/, {
                                data: data,
                                definition: definition,
                            }];
                }
            });
        });
    };
    return _RoleDefinition;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableInstance"]));

var RoleDefinition = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_RoleDefinition);
var PermissionKind;
(function (PermissionKind) {
    /**
     * Has no permissions on the Site. Not available through the user interface.
     */
    PermissionKind[PermissionKind["EmptyMask"] = 0] = "EmptyMask";
    /**
     * View items in lists, documents in document libraries, and Web discussion comments.
     */
    PermissionKind[PermissionKind["ViewListItems"] = 1] = "ViewListItems";
    /**
     * Add items to lists, documents to document libraries, and Web discussion comments.
     */
    PermissionKind[PermissionKind["AddListItems"] = 2] = "AddListItems";
    /**
     * Edit items in lists, edit documents in document libraries, edit Web discussion comments
     * in documents, and customize Web Part Pages in document libraries.
     */
    PermissionKind[PermissionKind["EditListItems"] = 3] = "EditListItems";
    /**
     * Delete items from a list, documents from a document library, and Web discussion
     * comments in documents.
     */
    PermissionKind[PermissionKind["DeleteListItems"] = 4] = "DeleteListItems";
    /**
     * Approve a minor version of a list item or document.
     */
    PermissionKind[PermissionKind["ApproveItems"] = 5] = "ApproveItems";
    /**
     * View the source of documents with server-side file handlers.
     */
    PermissionKind[PermissionKind["OpenItems"] = 6] = "OpenItems";
    /**
     * View past versions of a list item or document.
     */
    PermissionKind[PermissionKind["ViewVersions"] = 7] = "ViewVersions";
    /**
     * Delete past versions of a list item or document.
     */
    PermissionKind[PermissionKind["DeleteVersions"] = 8] = "DeleteVersions";
    /**
     * Discard or check in a document which is checked out to another user.
     */
    PermissionKind[PermissionKind["CancelCheckout"] = 9] = "CancelCheckout";
    /**
     * Create, change, and delete personal views of lists.
     */
    PermissionKind[PermissionKind["ManagePersonalViews"] = 10] = "ManagePersonalViews";
    /**
     * Create and delete lists, add or remove columns in a list, and add or remove public views of a list.
     */
    PermissionKind[PermissionKind["ManageLists"] = 12] = "ManageLists";
    /**
     * View forms, views, and application pages, and enumerate lists.
     */
    PermissionKind[PermissionKind["ViewFormPages"] = 13] = "ViewFormPages";
    /**
     * Make content of a list or document library retrieveable for anonymous users through SharePoint search.
     * The list permissions in the site do not change.
     */
    PermissionKind[PermissionKind["AnonymousSearchAccessList"] = 14] = "AnonymousSearchAccessList";
    /**
     * Allow users to open a Site, list, or folder to access items inside that container.
     */
    PermissionKind[PermissionKind["Open"] = 17] = "Open";
    /**
     * View pages in a Site.
     */
    PermissionKind[PermissionKind["ViewPages"] = 18] = "ViewPages";
    /**
     * Add, change, or delete HTML pages or Web Part Pages, and edit the Site using
     * a Windows SharePoint Services compatible editor.
     */
    PermissionKind[PermissionKind["AddAndCustomizePages"] = 19] = "AddAndCustomizePages";
    /**
     * Apply a theme or borders to the entire Site.
     */
    PermissionKind[PermissionKind["ApplyThemeAndBorder"] = 20] = "ApplyThemeAndBorder";
    /**
     * Apply a style sheet (.css file) to the Site.
     */
    PermissionKind[PermissionKind["ApplyStyleSheets"] = 21] = "ApplyStyleSheets";
    /**
     * View reports on Site usage.
     */
    PermissionKind[PermissionKind["ViewUsageData"] = 22] = "ViewUsageData";
    /**
     * Create a Site using Self-Service Site Creation.
     */
    PermissionKind[PermissionKind["CreateSSCSite"] = 23] = "CreateSSCSite";
    /**
     * Create subsites such as team sites, Meeting Workspace sites, and Document Workspace sites.
     */
    PermissionKind[PermissionKind["ManageSubwebs"] = 24] = "ManageSubwebs";
    /**
     * Create a group of users that can be used anywhere within the site collection.
     */
    PermissionKind[PermissionKind["CreateGroups"] = 25] = "CreateGroups";
    /**
     * Create and change permission levels on the Site and assign permissions to users
     * and groups.
     */
    PermissionKind[PermissionKind["ManagePermissions"] = 26] = "ManagePermissions";
    /**
     * Enumerate files and folders in a Site using Microsoft Office SharePoint Designer
     * and WebDAV interfaces.
     */
    PermissionKind[PermissionKind["BrowseDirectories"] = 27] = "BrowseDirectories";
    /**
     * View information about users of the Site.
     */
    PermissionKind[PermissionKind["BrowseUserInfo"] = 28] = "BrowseUserInfo";
    /**
     * Add or remove personal Web Parts on a Web Part Page.
     */
    PermissionKind[PermissionKind["AddDelPrivateWebParts"] = 29] = "AddDelPrivateWebParts";
    /**
     * Update Web Parts to display personalized information.
     */
    PermissionKind[PermissionKind["UpdatePersonalWebParts"] = 30] = "UpdatePersonalWebParts";
    /**
     * Grant the ability to perform all administration tasks for the Site as well as
     * manage content, activate, deactivate, or edit properties of Site scoped Features
     * through the object model or through the user interface (UI). When granted on the
     * root Site of a Site Collection, activate, deactivate, or edit properties of
     * site collection scoped Features through the object model. To browse to the Site
     * Collection Features page and activate or deactivate Site Collection scoped Features
     * through the UI, you must be a Site Collection administrator.
     */
    PermissionKind[PermissionKind["ManageWeb"] = 31] = "ManageWeb";
    /**
     * Content of lists and document libraries in the Web site will be retrieveable for anonymous users through
     * SharePoint search if the list or document library has AnonymousSearchAccessList set.
     */
    PermissionKind[PermissionKind["AnonymousSearchAccessWebLists"] = 32] = "AnonymousSearchAccessWebLists";
    /**
     * Use features that launch client applications. Otherwise, users must work on documents
     * locally and upload changes.
     */
    PermissionKind[PermissionKind["UseClientIntegration"] = 37] = "UseClientIntegration";
    /**
     * Use SOAP, WebDAV, or Microsoft Office SharePoint Designer interfaces to access the Site.
     */
    PermissionKind[PermissionKind["UseRemoteAPIs"] = 38] = "UseRemoteAPIs";
    /**
     * Manage alerts for all users of the Site.
     */
    PermissionKind[PermissionKind["ManageAlerts"] = 39] = "ManageAlerts";
    /**
     * Create e-mail alerts.
     */
    PermissionKind[PermissionKind["CreateAlerts"] = 40] = "CreateAlerts";
    /**
     * Allows a user to change his or her user information, such as adding a picture.
     */
    PermissionKind[PermissionKind["EditMyUserInfo"] = 41] = "EditMyUserInfo";
    /**
     * Enumerate permissions on Site, list, folder, document, or list item.
     */
    PermissionKind[PermissionKind["EnumeratePermissions"] = 63] = "EnumeratePermissions";
    /**
     * Has all permissions on the Site. Not available through the user interface.
     */
    PermissionKind[PermissionKind["FullMask"] = 65] = "FullMask";
})(PermissionKind || (PermissionKind = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "1jZ6":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/sites/index.js ***!
  \*********************************************/
/*! exports provided: Site */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "ilAS");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Site", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Site"]; });




Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"].prototype, "site", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_1__["Site"])(this._baseUrl).configure(this._options);
    },
});
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "1mTr":
/*!***********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/types.js ***!
  \***********************************************************/
/*! exports provided: _UserCustomActions, UserCustomActions, _UserCustomAction, UserCustomAction, UserCustomActionRegistrationType, UserCustomActionScope */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_UserCustomActions", function() { return _UserCustomActions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "UserCustomActions", function() { return UserCustomActions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_UserCustomAction", function() { return _UserCustomAction; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "UserCustomAction", function() { return UserCustomAction; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "UserCustomActionRegistrationType", function() { return UserCustomActionRegistrationType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "UserCustomActionScope", function() { return UserCustomActionScope; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../utils/metadata */ "z067");








var _UserCustomActions = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_UserCustomActions, _super);
    function _UserCustomActions() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Returns the user custom action with the specified id
     *
     * @param id The GUID id of the user custom action to retrieve
     */
    _UserCustomActions.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(UserCustomAction(this).concat("('" + id + "')"), "ucas.getById");
    };
    /**
     * Creates a user custom action
     *
     * @param properties The information object of property names and values which define the new user custom action
     */
    _UserCustomActions.prototype.add = function (properties) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_3__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_7__["metadata"])("SP.UserCustomAction"), properties)))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                action: this.getById(data.Id),
                                data: data,
                            }];
                }
            });
        });
    };
    /**
     * Deletes all user custom actions in the collection
     */
    _UserCustomActions.prototype.clear = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(UserCustomActions, "clear"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("ucas.add")
    ], _UserCustomActions.prototype, "add", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("ucas.clear")
    ], _UserCustomActions.prototype, "clear", null);
    _UserCustomActions = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_4__["defaultPath"])("usercustomactions")
    ], _UserCustomActions);
    return _UserCustomActions;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var UserCustomActions = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_UserCustomActions);
var _UserCustomAction = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_UserCustomAction, _super);
    function _UserCustomAction() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["deleteable"])("uca");
        /**
        * Updates this user custom action with the supplied properties
        *
        * @param properties An information object of property names and values to update for this user custom action
        */
        _this.update = _this._update("SP.UserCustomAction", function (data) { return ({ data: data, action: _this }); });
        return _this;
    }
    return _UserCustomAction;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var UserCustomAction = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_UserCustomAction);
var UserCustomActionRegistrationType;
(function (UserCustomActionRegistrationType) {
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["None"] = 0] = "None";
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["List"] = 1] = "List";
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["ContentType"] = 2] = "ContentType";
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["ProgId"] = 3] = "ProgId";
    UserCustomActionRegistrationType[UserCustomActionRegistrationType["FileType"] = 4] = "FileType";
})(UserCustomActionRegistrationType || (UserCustomActionRegistrationType = {}));
var UserCustomActionScope;
(function (UserCustomActionScope) {
    UserCustomActionScope[UserCustomActionScope["Unknown"] = 0] = "Unknown";
    UserCustomActionScope[UserCustomActionScope["Site"] = 2] = "Site";
    UserCustomActionScope[UserCustomActionScope["Web"] = 3] = "Web";
    UserCustomActionScope[UserCustomActionScope["List"] = 4] = "List";
})(UserCustomActionScope || (UserCustomActionScope = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "2B8P":
/*!***********************************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/user-custom-actions.js ***!
  \***********************************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _user_custom_actions_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../user-custom-actions/types */ "1mTr");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./funcs */ "T2j4");


_user_custom_actions_types__WEBPACK_IMPORTED_MODULE_0__["_UserCustomAction"].prototype.titleResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_1__["getValueForUICultureBinder"])("titleResource");
_user_custom_actions_types__WEBPACK_IMPORTED_MODULE_0__["_UserCustomAction"].prototype.descriptionResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_1__["getValueForUICultureBinder"])("descriptionResource");
//# sourceMappingURL=user-custom-actions.js.map

/***/ }),

/***/ "32VY":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/views/types.js ***!
  \*********************************************/
/*! exports provided: _Views, Views, _View, View, _ViewFields, ViewFields, ViewScope */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Views", function() { return _Views; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Views", function() { return Views; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_View", function() { return _View; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "View", function() { return View; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_ViewFields", function() { return _ViewFields; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ViewFields", function() { return ViewFields; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ViewScope", function() { return ViewScope; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../telemetry */ "h3gW");







var _Views = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Views, _super);
    function _Views() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a view by guid id
     *
     * @param id The GUID id of the view
     */
    _Views.prototype.getById = function (id) {
        return View(this).concat("('" + id + "')");
    };
    /**
     * Gets a view by title (case-sensitive)
     *
     * @param title The case-sensitive title of the view
     */
    _Views.prototype.getByTitle = function (title) {
        return View(this, "getByTitle('" + title + "')");
    };
    /**
     * Adds a new view to the collection
     *
     * @param title The new views's title
     * @param personalView True if this is a personal view, otherwise false, default = false
     * @param additionalSettings Will be passed as part of the view creation body
     */
    _Views.prototype.add = function (title, personalView, additionalSettings) {
        if (personalView === void 0) { personalView = false; }
        if (additionalSettings === void 0) { additionalSettings = {}; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(Object.assign(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])("SP.View"), {
                            "PersonalView": personalView,
                            "Title": title,
                        }, additionalSettings));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(Views, null), postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                view: this.getById(data.Id),
                            }];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("vs.add")
    ], _Views.prototype, "add", null);
    _Views = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_4__["defaultPath"])("views")
    ], _Views);
    return _Views;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var Views = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_Views);
var _View = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_View, _super);
    function _View() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["deleteable"])("vw");
        /**
         * Updates this view intance with the supplied properties
         *
         * @param properties A plain object hash of values to update for the view
         */
        _this.update = _this._update("SP.View", function (data) { return ({ data: data, view: _this }); });
        return _this;
    }
    Object.defineProperty(_View.prototype, "fields", {
        get: function () {
            return ViewFields(this);
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Returns the list view as HTML.
     *
     */
    _View.prototype.renderAsHtml = function () {
        return this.clone(View, "renderashtml")();
    };
    /**
     * Sets the view schema
     *
     * @param viewXml The view XML to set
     */
    _View.prototype.setViewXml = function (viewXml) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(View, "SetViewXml"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ viewXml: viewXml }));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("v.renderAsHtml")
    ], _View.prototype, "renderAsHtml", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("v.setViewXml")
    ], _View.prototype, "setViewXml", null);
    return _View;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableInstance"]));

var View = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_View);
var _ViewFields = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_ViewFields, _super);
    function _ViewFields() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a value that specifies the XML schema that represents the collection.
     */
    _ViewFields.prototype.getSchemaXml = function () {
        return this.clone(ViewFields, "schemaxml")();
    };
    /**
     * Adds the field with the specified field internal name or display name to the collection.
     *
     * @param fieldTitleOrInternalName The case-sensitive internal name or display name of the field to add.
     */
    _ViewFields.prototype.add = function (fieldTitleOrInternalName) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(ViewFields, "addviewfield('" + fieldTitleOrInternalName + "')"));
    };
    /**
     * Moves the field with the specified field internal name to the specified position in the collection.
     *
     * @param field The case-sensitive internal name of the field to move.
     * @param index The zero-based index of the new position for the field.
     */
    _ViewFields.prototype.move = function (field, index) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(ViewFields, "moveviewfieldto"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ field: field, index: index }));
    };
    /**
     * Removes all the fields from the collection.
     */
    _ViewFields.prototype.removeAll = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(ViewFields, "removeallviewfields"));
    };
    /**
     * Removes the field with the specified field internal name from the collection.
     *
     * @param fieldInternalName The case-sensitive internal name of the field to remove from the view.
     */
    _ViewFields.prototype.remove = function (fieldInternalName) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(ViewFields, "removeviewfield('" + fieldInternalName + "')"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("vfs.getSchemaXml")
    ], _ViewFields.prototype, "getSchemaXml", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("vfs.add")
    ], _ViewFields.prototype, "add", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("vfs.move")
    ], _ViewFields.prototype, "move", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("vfs.removeAll")
    ], _ViewFields.prototype, "removeAll", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("vfs.remove")
    ], _ViewFields.prototype, "remove", null);
    _ViewFields = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_4__["defaultPath"])("viewfields")
    ], _ViewFields);
    return _ViewFields;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var ViewFields = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_ViewFields);
var ViewScope;
(function (ViewScope) {
    ViewScope[ViewScope["DefaultValue"] = 0] = "DefaultValue";
    ViewScope[ViewScope["Recursive"] = 1] = "Recursive";
    ViewScope[ViewScope["RecursiveAll"] = 2] = "RecursiveAll";
    ViewScope[ViewScope["FilesOnly"] = 3] = "FilesOnly";
})(ViewScope || (ViewScope = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "36LU":
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/clientside-pages/index.js ***!
  \********************************************************/
/*! exports provided: ClientsidePageFromFile, CreateClientsidePage, CanvasColumn, CanvasSection, ClientsideText, ClientsideWebpart, ColumnControl, PromotedState */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "laIm");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "UPqE");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ClientsidePageFromFile", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ClientsidePageFromFile"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CreateClientsidePage", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["CreateClientsidePage"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CanvasColumn", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["CanvasColumn"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CanvasSection", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["CanvasSection"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ClientsideText", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ClientsideText"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ClientsideWebpart", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ClientsideWebpart"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ColumnControl", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ColumnControl"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PromotedState", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["PromotedState"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "37pO":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/hubsites/web.js ***!
  \**********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../operations */ "UK2s");



_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.hubSiteData = function (forceRefresh) {
    if (forceRefresh === void 0) { forceRefresh = false; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var data;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, this.clone(_webs_types__WEBPACK_IMPORTED_MODULE_1__["Web"], "hubSiteData(" + forceRefresh + ")")()];
                case 1:
                    data = _a.sent();
                    if (typeof data === "string") {
                        return [2 /*return*/, JSON.parse(data)];
                    }
                    return [2 /*return*/, data];
            }
        });
    });
};
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.syncHubSiteTheme = function () {
    return Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this.clone(_webs_types__WEBPACK_IMPORTED_MODULE_1__["Web"], "syncHubSiteTheme"));
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "3DT9":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/items/types.js ***!
  \*********************************************/
/*! exports provided: _Items, Items, _Item, Item, _ItemVersions, ItemVersions, _ItemVersion, ItemVersion, PagedItemCollection */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Items", function() { return _Items; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Items", function() { return Items; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Item", function() { return _Item; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Item", function() { return Item; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_ItemVersions", function() { return _ItemVersions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ItemVersions", function() { return ItemVersions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_ItemVersion", function() { return _ItemVersion; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ItemVersion", function() { return ItemVersion; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PagedItemCollection", function() { return PagedItemCollection; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _pnp_logging__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/logging */ "B3L+");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../telemetry */ "h3gW");










/**
 * Describes a collection of Item objects
 *
 */
var _Items = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Items, _super);
    function _Items() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
    * Gets an Item by id
    *
    * @param id The integer id of the item to retrieve
    */
    _Items.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Item(this).concat("(" + id + ")"), "is.getById");
    };
    /**
     * Gets BCS Item by string id
     *
     * @param stringId The string id of the BCS item to retrieve
     */
    _Items.prototype.getItemByStringId = function (stringId) {
        // creates an item with the parent list path and append out method call
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Item(this.parentUrl, "getItemByStringId('" + stringId + "')"), "is.getItemByStringId");
    };
    /**
     * Skips the specified number of items (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#sectionSection6)
     *
     * @param skip The starting id where the page should start, use with top to specify pages
     * @param reverse It true the PagedPrev=true parameter is added allowing backwards navigation in the collection
     */
    _Items.prototype.skip = function (skip, reverse) {
        if (reverse === void 0) { reverse = false; }
        if (reverse) {
            this.query.set("$skiptoken", encodeURIComponent("Paged=TRUE&PagedPrev=TRUE&p_ID=" + skip));
        }
        else {
            this.query.set("$skiptoken", encodeURIComponent("Paged=TRUE&p_ID=" + skip));
        }
        return this;
    };
    /**
     * Gets a collection designed to aid in paging through data
     *
     */
    _Items.prototype.getPaged = function () {
        return this.usingParser(new PagedItemCollectionParser(this))();
    };
    /**
     * Gets all the items in a list, regardless of count. Does not support batching or caching
     *
     *  @param requestSize Number of items to return in each request (Default: 2000)
     *  @param acceptHeader Allows for setting the value of the Accept header for SP 2013 support
     */
    _Items.prototype.getAll = function (requestSize, acceptHeader) {
        if (requestSize === void 0) { requestSize = 2000; }
        if (acceptHeader === void 0) { acceptHeader = "application/json;odata=nometadata"; }
        _pnp_logging__WEBPACK_IMPORTED_MODULE_5__["Logger"].write("Calling items.getAll should be done sparingly. Ensure this is the correct choice. If you are unsure, it is not.", 2 /* Warning */);
        // this will be used for the actual query
        // and we set no metadata here to try and reduce traffic
        var items = Items(this, "").top(requestSize).configure({
            headers: {
                "Accept": acceptHeader,
            },
        });
        // let's copy over the odata query params that can be applied
        // $top - allow setting the page size this way (override what we did above)
        // $select - allow picking the return fields (good behavior)
        // $filter - allow setting a filter, though this may fail due for large lists
        this.query.forEach(function (v, k) {
            if (/^\$select|filter|top|expand$/i.test(k)) {
                items.query.set(k, v);
            }
        });
        // give back the promise
        return new Promise(function (resolve, reject) {
            // this will eventually hold the items we return
            var itemsCollector = [];
            // action that will gather up our results recursively
            var gatherer = function (last) {
                // collect that set of results
                [].push.apply(itemsCollector, last.results);
                // if we have more, repeat - otherwise resolve with the collected items
                if (last.hasNext) {
                    last.getNext().then(gatherer).catch(reject);
                }
                else {
                    resolve(itemsCollector);
                }
            };
            // start the cycle
            items.getPaged().then(gatherer).catch(reject);
        });
    };
    /**
     * Adds a new item to the collection
     *
     * @param properties The new items's properties
     * @param listItemEntityTypeFullName The type name of the list's entities
     */
    _Items.prototype.add = function (properties, listItemEntityTypeFullName) {
        if (properties === void 0) { properties = {}; }
        if (listItemEntityTypeFullName === void 0) { listItemEntityTypeFullName = null; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var removeDependency, listItemEntityType, postBody, promise;
            var _this = this;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        removeDependency = this.addBatchDependency();
                        return [4 /*yield*/, this.ensureListItemEntityTypeName(listItemEntityTypeFullName)];
                    case 1:
                        listItemEntityType = _a.sent();
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_6__["metadata"])(listItemEntityType), properties));
                        promise = Object(_operations__WEBPACK_IMPORTED_MODULE_8__["spPost"])(this.clone(Items, ""), postBody).then(function (data) {
                            return {
                                data: data,
                                item: _this.getById(data.Id),
                            };
                        });
                        removeDependency();
                        return [2 /*return*/, promise];
                }
            });
        });
    };
    /**
     * Ensures we have the proper list item entity type name, either from the value provided or from the list
     *
     * @param candidatelistItemEntityTypeFullName The potential type name
     */
    _Items.prototype.ensureListItemEntityTypeName = function (candidatelistItemEntityTypeFullName) {
        return candidatelistItemEntityTypeFullName ?
            Promise.resolve(candidatelistItemEntityTypeFullName) :
            this.getParent(_lists_types__WEBPACK_IMPORTED_MODULE_3__["List"]).getListItemEntityTypeFullName();
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("is.getPaged")
    ], _Items.prototype, "getPaged", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("is.getAll")
    ], _Items.prototype, "getAll", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("is.add")
    ], _Items.prototype, "add", null);
    _Items = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_7__["defaultPath"])("items")
    ], _Items);
    return _Items;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var Items = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Items);
/**
 * Descrines a single Item instance
 *
 */
var _Item = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Item, _super);
    function _Item() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["deleteableWithETag"])("i");
        return _this;
    }
    Object.defineProperty(_Item.prototype, "effectiveBasePermissions", {
        /**
         * Gets the effective base permissions for the item
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryable"])(this, "EffectiveBasePermissions"), "i.effectiveBasePermissions");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Item.prototype, "effectiveBasePermissionsForUI", {
        /**
         * Gets the effective base permissions for the item in a UI context
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryable"])(this, "EffectiveBasePermissionsForUI"), "i.effectiveBasePermissionsForUI");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Item.prototype, "fieldValuesAsHTML", {
        /**
         * Gets the field values for this list item in their HTML representation
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryableInstance"])(this, "FieldValuesAsHTML"), "i.fvHTML");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Item.prototype, "fieldValuesAsText", {
        /**
         * Gets the field values for this list item in their text representation
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryableInstance"])(this, "FieldValuesAsText"), "i.fvText");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Item.prototype, "fieldValuesForEdit", {
        /**
         * Gets the field values for this list item for use in editing controls
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryableInstance"])(this, "FieldValuesForEdit"), "i.fvEdit");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Item.prototype, "versions", {
        /**
         * Gets the collection of versions associated with this item
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(ItemVersions(this), "i.versions");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Item.prototype, "list", {
        get: function () {
            return this.getParent(_lists_types__WEBPACK_IMPORTED_MODULE_3__["List"], this.parentUrl.substr(0, this.parentUrl.lastIndexOf("/")));
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Updates this list intance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the list
     * @param eTag Value used in the IF-Match header, by default "*"
     * @param listItemEntityTypeFullName The type name of the list's entities
     */
    _Item.prototype.update = function (properties, eTag, listItemEntityTypeFullName) {
        if (eTag === void 0) { eTag = "*"; }
        if (listItemEntityTypeFullName === void 0) { listItemEntityTypeFullName = null; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var removeDependency, listItemEntityType, postBody, poster, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        removeDependency = this.addBatchDependency();
                        return [4 /*yield*/, this.ensureListItemEntityTypeName(listItemEntityTypeFullName)];
                    case 1:
                        listItemEntityType = _a.sent();
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_6__["metadata"])(listItemEntityType), properties), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["headers"])({
                            "IF-Match": eTag,
                            "X-HTTP-Method": "MERGE",
                        }));
                        removeDependency();
                        poster = _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(this.clone(Item).usingParser(new ItemUpdatedParser()), "i.update");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_8__["spPost"])(poster, postBody)];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                item: this,
                            }];
                }
            });
        });
    };
    /**
     * Moves the list item to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     */
    _Item.prototype.recycle = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_8__["spPost"])(this.clone(Item, "recycle"));
    };
    /**
     * Gets a string representation of the full URL to the WOPI frame.
     * If there is no associated WOPI application, or no associated action, an empty string is returned.
     *
     * @param action Display mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
     */
    _Item.prototype.getWopiFrameUrl = function (action) {
        if (action === void 0) { action = 0; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var i, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        i = this.clone(Item, "getWOPIFrameUrl(@action)");
                        i.query.set("@action", action);
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_8__["spPost"])(i)];
                    case 1:
                        data = _a.sent();
                        // handle verbose mode
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(data, "GetWOPIFrameUrl")) {
                            return [2 /*return*/, data.GetWOPIFrameUrl];
                        }
                        return [2 /*return*/, data];
                }
            });
        });
    };
    /**
     * Validates and sets the values of the specified collection of fields for the list item.
     *
     * @param formValues The fields to change and their new values.
     * @param bNewDocumentUpdate true if the list item is a document being updated after upload; otherwise false.
     */
    _Item.prototype.validateUpdateListItem = function (formValues, bNewDocumentUpdate) {
        if (bNewDocumentUpdate === void 0) { bNewDocumentUpdate = false; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_8__["spPost"])(this.clone(Item, "validateupdatelistitem"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({ formValues: formValues, bNewDocumentUpdate: bNewDocumentUpdate }));
    };
    /**
     * Ensures we have the proper list item entity type name, either from the value provided or from the list
     *
     * @param candidatelistItemEntityTypeFullName The potential type name
     */
    _Item.prototype.ensureListItemEntityTypeName = function (candidatelistItemEntityTypeFullName) {
        return candidatelistItemEntityTypeFullName ?
            Promise.resolve(candidatelistItemEntityTypeFullName) :
            this.list.getListItemEntityTypeFullName();
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("i.recycle")
    ], _Item.prototype, "recycle", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("i.getWopiFrameUrl")
    ], _Item.prototype, "getWopiFrameUrl", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("i.validateUpdateListItem")
    ], _Item.prototype, "validateUpdateListItem", null);
    return _Item;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Item = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Item);
/**
 * Describes a collection of Version objects
 *
 */
var _ItemVersions = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_ItemVersions, _super);
    function _ItemVersions() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a version by id
     *
     * @param versionId The id of the version to retrieve
     */
    _ItemVersions.prototype.getById = function (versionId) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(ItemVersion(this).concat("(" + versionId + ")"), "iv.getById");
    };
    _ItemVersions = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_7__["defaultPath"])("versions")
    ], _ItemVersions);
    return _ItemVersions;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var ItemVersions = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_ItemVersions);
/**
 * Describes a single Version instance
 *
 */
var _ItemVersion = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_ItemVersion, _super);
    function _ItemVersion() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["deleteableWithETag"])("iv");
        return _this;
    }
    return _ItemVersion;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var ItemVersion = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_ItemVersion);
/**
 * Provides paging functionality for list items
 */
var PagedItemCollection = /** @class */ (function () {
    function PagedItemCollection(parent, nextUrl, results) {
        this.parent = parent;
        this.nextUrl = nextUrl;
        this.results = results;
    }
    Object.defineProperty(PagedItemCollection.prototype, "hasNext", {
        /**
         * If true there are more results available in the set, otherwise there are not
         */
        get: function () {
            return typeof this.nextUrl === "string" && this.nextUrl.length > 0;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Gets the next set of results, or resolves to null if no results are available
     */
    PagedItemCollection.prototype.getNext = function () {
        if (this.hasNext) {
            var items = _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Items(this.nextUrl, null).configureFrom(this.parent), "ip.getNext");
            return items.getPaged();
        }
        return new Promise(function (r) { return r(null); });
    };
    return PagedItemCollection;
}());

var PagedItemCollectionParser = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(PagedItemCollectionParser, _super);
    function PagedItemCollectionParser(_parent) {
        var _this = _super.call(this) || this;
        _this._parent = _parent;
        return _this;
    }
    PagedItemCollectionParser.prototype.parse = function (r) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (_this.handleError(r, reject)) {
                r.json().then(function (json) {
                    var nextUrl = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(json, "d") && Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(json.d, "__next") ? json.d.__next : json["odata.nextLink"];
                    resolve(new PagedItemCollection(_this._parent, nextUrl, _this.parseODataJSON(json)));
                });
            }
        });
    };
    return PagedItemCollectionParser;
}(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["ODataParser"]));
var ItemUpdatedParser = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(ItemUpdatedParser, _super);
    function ItemUpdatedParser() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ItemUpdatedParser.prototype.parse = function (r) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (_this.handleError(r, reject)) {
                resolve({
                    "odata.etag": r.headers.get("etag"),
                });
            }
        });
    };
    return ItemUpdatedParser;
}(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["ODataParser"]));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "3fSn":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/forms/list.js ***!
  \********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "zYJ9");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "forms", _types__WEBPACK_IMPORTED_MODULE_2__["Forms"], "forms");
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "4Egx":
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/item.js ***!
  \****************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "S3gC");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"], "contentType", _types__WEBPACK_IMPORTED_MODULE_2__["ContentType"], "ContentType");
//# sourceMappingURL=item.js.map

/***/ }),

/***/ "4G66":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/security/web.js ***!
  \**********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "1dmT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./funcs */ "vXSC");





Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "roleDefinitions", _types__WEBPACK_IMPORTED_MODULE_2__["RoleDefinitions"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "roleAssignments", _types__WEBPACK_IMPORTED_MODULE_2__["RoleAssignments"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "firstUniqueAncestorSecurableObject", _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"]);
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.getUserEffectivePermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["getUserEffectivePermissions"];
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.getCurrentUserEffectivePermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["getCurrentUserEffectivePermissions"];
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.breakRoleInheritance = _funcs__WEBPACK_IMPORTED_MODULE_4__["breakRoleInheritance"];
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.resetRoleInheritance = _funcs__WEBPACK_IMPORTED_MODULE_4__["resetRoleInheritance"];
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.userHasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["userHasPermissions"];
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.currentUserHasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["currentUserHasPermissions"];
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.hasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["hasPermissions"];
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "4tXH":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/types.js ***!
  \***********************************************/
/*! exports provided: _Folders, Folders, _Folder, Folder */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Folders", function() { return _Folders; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Folders", function() { return Folders; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Folder", function() { return _Folder; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Folder", function() { return Folder; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../utils/toResourcePath */ "t8hP");












var _Folders = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Folders, _super);
    function _Folders() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a folder by it's name
     *
     * @param name Folder's name
     */
    _Folders.prototype.getByName = function (name) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Folder(this).concat("('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(name) + "')"), "fs.getByName");
    };
    /**
     * Adds a new folder at the specified URL
     *
     * @param url
     */
    _Folders.prototype.add = function (url) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Folders, "add('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(url) + "')"))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                folder: this.getByName(url),
                            }];
                }
            });
        });
    };
    /**
     * Adds a new folder by path and should be prefered over add
     *
     * @param serverRelativeUrl The server relative url of the new folder to create
     * @param overwrite True to overwrite an existing folder, default false
     */
    _Folders.prototype.addUsingPath = function (serverRelativeUrl, overwrite) {
        if (overwrite === void 0) { overwrite = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Folders, "addUsingPath(DecodedUrl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(serverRelativeUrl) + "',overwrite=" + overwrite + ")"))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                folder: Folder(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(this.toUrl()), "_api/web/getFolderByServerRelativePath(decodedUrl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(serverRelativeUrl) + "')"),
                            }];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fs.add")
    ], _Folders.prototype, "add", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fs.addUsingPath")
    ], _Folders.prototype, "addUsingPath", null);
    _Folders = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("folders")
    ], _Folders);
    return _Folders;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var Folders = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_Folders);
var _Folder = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Folder, _super);
    function _Folder() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["deleteableWithETag"])("f");
        /**
         * Updates folder's properties
         * @param props Folder's properties to update
         */
        _this.update = _this._update("SP.Folder", function (data) { return ({ data: data, folder: _this }); });
        return _this;
    }
    Object.defineProperty(_Folder.prototype, "contentTypeOrder", {
        /**
         * Specifies the sequence in which content types are displayed.
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableCollection"])(this, "contentTypeOrder"), "f.contentTypeOrder");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Folder.prototype, "folders", {
        /**
         * Gets this folder's sub folders
         *
         */
        get: function () {
            return Folders(this);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Folder.prototype, "listItemAllFields", {
        /**
         * Gets this folder's list item field values
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableInstance"])(this, "listItemAllFields"), "f.listItemAllFields");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Folder.prototype, "parentFolder", {
        /**
         * Gets the parent folder, if available
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Folder(this, "parentFolder"), "f.parentFolder");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Folder.prototype, "properties", {
        /**
         * Gets this folder's properties
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableInstance"])(this, "properties"), "f.properties");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Folder.prototype, "serverRelativeUrl", {
        /**
         * Gets this folder's server relative url
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryable"])(this, "serverRelativeUrl"), "f.serverRelativeUrl");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Folder.prototype, "uniqueContentTypeOrder", {
        /**
         * Gets a value that specifies the content type order.
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableCollection"])(this, "uniqueContentTypeOrder"), "f.uniqueContentTypeOrder");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Moves the folder to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     */
    _Folder.prototype.recycle = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Folder, "recycle"));
    };
    /**
     * Gets the associated list item for this folder, loading the default properties
     */
    _Folder.prototype.getItem = function () {
        var selects = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            selects[_i] = arguments[_i];
        }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var q;
            var _a;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, (_a = this.listItemAllFields).select.apply(_a, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])(selects))()];
                    case 1:
                        q = _b.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_items_types__WEBPACK_IMPORTED_MODULE_4__["Item"])(Object(_odata__WEBPACK_IMPORTED_MODULE_3__["odataUrlFrom"])(q)), q)];
                }
            });
        });
    };
    /**
     * Moves a folder to destination path
     *
     * @param destUrl Absolute or relative URL of the destination path
     */
    _Folder.prototype.moveTo = function (destUrl) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var _a, srcUrl, absoluteUrl, webBaseUrl, hostUrl;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.select("ServerRelativeUrl")()];
                    case 1:
                        _a = _b.sent(), srcUrl = _a.ServerRelativeUrl, absoluteUrl = _a["odata.id"];
                        webBaseUrl = Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(absoluteUrl);
                        hostUrl = webBaseUrl.replace("://", "___").split("/")[0].replace("___", "://");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Folder(webBaseUrl, "/_api/SP.MoveCopyUtil.MoveFolder()"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
                                destUrl: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isUrlAbsolute"])(destUrl) ? destUrl : "" + hostUrl + destUrl,
                                srcUrl: "" + hostUrl + srcUrl,
                            }))];
                    case 2:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Moves a folder by path to destination path
     * Also works with different site collections.
     *
     * @param destUrl Absolute or relative URL of the destination path
     * @param keepBoth Keep both if folder with the same name in the same location already exists?
     */
    _Folder.prototype.moveByPath = function (destUrl, KeepBoth) {
        if (KeepBoth === void 0) { KeepBoth = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var _a, srcUrl, absoluteUrl, webBaseUrl, hostUrl;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.select("ServerRelativeUrl")()];
                    case 1:
                        _a = _b.sent(), srcUrl = _a.ServerRelativeUrl, absoluteUrl = _a["odata.id"];
                        webBaseUrl = Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(absoluteUrl);
                        hostUrl = webBaseUrl.replace("://", "___").split("/")[0].replace("___", "://");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Folder(webBaseUrl, "/_api/SP.MoveCopyUtil.MoveFolderByPath()"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
                                destPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__["toResourcePath"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isUrlAbsolute"])(destUrl) ? destUrl : "" + hostUrl + destUrl),
                                options: {
                                    KeepBoth: KeepBoth,
                                    ResetAuthorAndCreatedOnCopy: true,
                                    ShouldBypassSharedLocks: true,
                                    __metadata: {
                                        type: "SP.MoveCopyOptions",
                                    },
                                },
                                srcPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__["toResourcePath"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isUrlAbsolute"])(srcUrl) ? srcUrl : "" + hostUrl + srcUrl),
                            }))];
                    case 2:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Copies a folder to destination path
     *
     * @param destUrl Absolute or relative URL of the destination path
     */
    _Folder.prototype.copyTo = function (destUrl) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var _a, srcUrl, absoluteUrl, webBaseUrl, hostUrl;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.select("ServerRelativeUrl")()];
                    case 1:
                        _a = _b.sent(), srcUrl = _a.ServerRelativeUrl, absoluteUrl = _a["odata.id"];
                        webBaseUrl = Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(absoluteUrl);
                        hostUrl = webBaseUrl.replace("://", "___").split("/")[0].replace("___", "://");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Folder(webBaseUrl, "/_api/SP.MoveCopyUtil.CopyFolder()"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
                                destUrl: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isUrlAbsolute"])(destUrl) ? destUrl : "" + hostUrl + destUrl,
                                srcUrl: "" + hostUrl + srcUrl,
                            }))];
                    case 2:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Copies a folder by path to destination path
     * Also works with different site collections.
     *
     * @param destUrl Absolute or relative URL of the destination path
     * @param keepBoth Keep both if folder with the same name in the same location already exists?
     */
    _Folder.prototype.copyByPath = function (destUrl, KeepBoth) {
        if (KeepBoth === void 0) { KeepBoth = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var _a, srcUrl, absoluteUrl, webBaseUrl, hostUrl;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.select("ServerRelativeUrl")()];
                    case 1:
                        _a = _b.sent(), srcUrl = _a.ServerRelativeUrl, absoluteUrl = _a["odata.id"];
                        webBaseUrl = Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(absoluteUrl);
                        hostUrl = webBaseUrl.replace("://", "___").split("/")[0].replace("___", "://");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Folder(webBaseUrl, "/_api/SP.MoveCopyUtil.CopyFolderByPath()"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
                                destPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__["toResourcePath"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isUrlAbsolute"])(destUrl) ? destUrl : "" + hostUrl + destUrl),
                                options: {
                                    KeepBoth: KeepBoth,
                                    ResetAuthorAndCreatedOnCopy: true,
                                    ShouldBypassSharedLocks: true,
                                    __metadata: {
                                        type: "SP.MoveCopyOptions",
                                    },
                                },
                                srcPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__["toResourcePath"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isUrlAbsolute"])(srcUrl) ? srcUrl : "" + hostUrl + srcUrl),
                            }))];
                    case 2:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Gets the shareable item associated with this folder
     */
    _Folder.prototype.getShareable = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var d, shareable;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableInstance"], "listItemAllFields", false).select("odata.id")()];
                    case 1:
                        d = _a.sent();
                        shareable = Object(_items_types__WEBPACK_IMPORTED_MODULE_4__["Item"])(Object(_odata__WEBPACK_IMPORTED_MODULE_3__["odataUrlFrom"])(d));
                        // we need to handle batching
                        if (this.hasBatch) {
                            shareable = shareable.inBatch(this.batch);
                        }
                        return [2 /*return*/, shareable];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("f.recycle")
    ], _Folder.prototype, "recycle", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("f.getItem")
    ], _Folder.prototype, "getItem", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("f.moveTo")
    ], _Folder.prototype, "moveTo", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("f.moveByPath")
    ], _Folder.prototype, "moveByPath", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("f.copyTo")
    ], _Folder.prototype, "copyTo", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("f.copyByPath")
    ], _Folder.prototype, "copyByPath", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("f.getShareable")
    ], _Folder.prototype, "getShareable", null);
    return _Folder;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableInstance"]));

var Folder = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_Folder);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "5dQz":
/*!***************************************!*\
  !*** ./node_modules/@pnp/sp/odata.js ***!
  \***************************************/
/*! exports provided: odataUrlFrom, spODataEntity, spODataEntityArray */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "odataUrlFrom", function() { return odataUrlFrom; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spODataEntity", function() { return spODataEntity; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spODataEntityArray", function() { return spODataEntityArray; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_logging__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/logging */ "B3L+");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./utils/extractweburl */ "tSEd");





function odataUrlFrom(candidate) {
    var parts = [];
    var s = ["odata.type", "odata.editLink", "__metadata", "odata.metadata", "odata.id"];
    if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(candidate, s[0]) && candidate[s[0]] === "SP.Web") {
        // webs return an absolute url in the id
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(candidate, s[4])) {
            parts.push(candidate[s[4]]);
        }
        else if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(candidate, s[2])) {
            // we are dealing with verbose, which has an absolute uri
            parts.push(candidate.__metadata.uri);
        }
    }
    else {
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(candidate, s[3]) && Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(candidate, s[1])) {
            // we are dealign with minimal metadata (default)
            // some entities return an abosolute url in the editlink while for others it is relative
            // without the _api. This code is meant to handle both situations
            var editLink = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isUrlAbsolute"])(candidate[s[1]]) ? candidate[s[1]].split("_api")[1] : candidate[s[1]];
            parts.push(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__["extractWebUrl"])(candidate[s[3]]), "_api", editLink);
        }
        else if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(candidate, s[1])) {
            parts.push("_api", candidate[s[1]]);
        }
        else if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(candidate, s[2])) {
            // we are dealing with verbose, which has an absolute uri
            parts.push(candidate.__metadata.uri);
        }
    }
    if (parts.length < 1) {
        _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].write("No uri information found in ODataEntity parsing, chaining will fail for this object.", 2 /* Warning */);
        return "";
    }
    return _pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"].apply(void 0, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])(parts));
}
var SPODataEntityParserImpl = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(SPODataEntityParserImpl, _super);
    function SPODataEntityParserImpl(factory) {
        var _this = _super.call(this) || this;
        _this.factory = factory;
        _this.hydrate = function (d) {
            var o = _this.factory(odataUrlFrom(d), null);
            return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(o, d);
        };
        return _this;
    }
    SPODataEntityParserImpl.prototype.parse = function (r) {
        var _this = this;
        return _super.prototype.parse.call(this, r).then(function (d) {
            var o = _this.factory(odataUrlFrom(d), null);
            return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(o, d);
        });
    };
    return SPODataEntityParserImpl;
}(_pnp_odata__WEBPACK_IMPORTED_MODULE_3__["ODataParser"]));
var SPODataEntityArrayParserImpl = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(SPODataEntityArrayParserImpl, _super);
    function SPODataEntityArrayParserImpl(factory) {
        var _this = _super.call(this) || this;
        _this.factory = factory;
        _this.hydrate = function (d) {
            return d.map(function (v) {
                var o = _this.factory(odataUrlFrom(v), null);
                return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(o, v);
            });
        };
        return _this;
    }
    SPODataEntityArrayParserImpl.prototype.parse = function (r) {
        var _this = this;
        return _super.prototype.parse.call(this, r).then(function (d) {
            return d.map(function (v) {
                var o = _this.factory(odataUrlFrom(v), null);
                return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(o, v);
            });
        });
    };
    return SPODataEntityArrayParserImpl;
}(_pnp_odata__WEBPACK_IMPORTED_MODULE_3__["ODataParser"]));
function spODataEntity(factory) {
    return new SPODataEntityParserImpl(factory);
}
function spODataEntityArray(factory) {
    return new SPODataEntityArrayParserImpl(factory);
}
//# sourceMappingURL=odata.js.map

/***/ }),

/***/ "5e8D":
/*!****************************************************!*\
  !*** ./node_modules/@pnp/odata/pipeline-binder.js ***!
  \****************************************************/
/*! exports provided: pipelineBinder, defaultPipelineBinder */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "pipelineBinder", function() { return pipelineBinder; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "defaultPipelineBinder", function() { return defaultPipelineBinder; });
/* harmony import */ var _parsers__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./parsers */ "xbXq");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _queryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./queryable */ "dKLh");
/* harmony import */ var _pipeline__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./pipeline */ "TZA7");




// first we bind the pipeline we will use for all requests within this closure
function pipelineBinder(pipes) {
    // then we bind the client factory we'll use (typically done in an implementing library such as sp)
    return function (clientFactory) {
        // then we create a binder we can apply for each type of method (GET, POST, etc.)
        return function (method) {
            // finally we get a function back to which we can pass an IQueryableData instance and execute the request it defines
            return function (o) {
                // send the IQueryableData down the pipeline
                return Object(_pipeline__WEBPACK_IMPORTED_MODULE_3__["pipe"])(Object.assign({}, {
                    batch: null,
                    batchDependency: null,
                    batchIndex: -1,
                    cachingOptions: null,
                    clientFactory: clientFactory,
                    cloneParentCacheOptions: null,
                    cloneParentWasCaching: false,
                    hasResult: false,
                    isBatched: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["objectDefinedNotNull"])(o.batch),
                    method: method,
                    options: null,
                    parentUrl: "",
                    parser: new _parsers__WEBPACK_IMPORTED_MODULE_0__["ODataParser"](),
                    pipes: pipes.slice(0),
                    query: new Map(),
                    requestId: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["getGUID"])(),
                    url: "",
                    useCaching: /^get$/i.test(o.method) && o.useCaching,
                }, Object(_queryable__WEBPACK_IMPORTED_MODULE_2__["cloneQueryableData"])(o)));
            };
        };
    };
}
var defaultPipelineBinder = pipelineBinder(Object(_pipeline__WEBPACK_IMPORTED_MODULE_3__["getDefaultPipeline"])());
//# sourceMappingURL=pipeline-binder.js.map

/***/ }),

/***/ "6a65":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/social/types.js ***!
  \**********************************************/
/*! exports provided: _Social, Social, _MySocial, MySocial, SocialActorType, SocialActorTypes, SocialFollowResult, SocialStatusCode */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Social", function() { return _Social; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Social", function() { return Social; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_MySocial", function() { return _MySocial; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "MySocial", function() { return MySocial; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SocialActorType", function() { return SocialActorType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SocialActorTypes", function() { return SocialActorTypes; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SocialFollowResult", function() { return SocialFollowResult; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SocialStatusCode", function() { return SocialStatusCode; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../telemetry */ "h3gW");








var _Social = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Social, _super);
    function _Social() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(_Social.prototype, "my", {
        get: function () {
            return MySocial(this);
        },
        enumerable: true,
        configurable: true
    });
    _Social.prototype.getFollowedSitesUri = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var r;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(SocialCloneFactory, "FollowedSitesUri").get()];
                    case 1:
                        r = _a.sent();
                        return [2 /*return*/, r.FollowedSitesUri || r];
                }
            });
        });
    };
    _Social.prototype.getFollowedDocumentsUri = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var r;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(SocialCloneFactory, "FollowedDocumentsUri").get()];
                    case 1:
                        r = _a.sent();
                        return [2 /*return*/, r.FollowedDocumentsUri || r];
                }
            });
        });
    };
    _Social.prototype.follow = function (actorInfo) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(SocialCloneFactory, "follow"), this.createSocialActorInfoRequestBody(actorInfo))];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    _Social.prototype.isFollowed = function (actorInfo) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(SocialCloneFactory, "isfollowed"), this.createSocialActorInfoRequestBody(actorInfo))];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    _Social.prototype.stopFollowing = function (actorInfo) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(SocialCloneFactory, "stopfollowing"), this.createSocialActorInfoRequestBody(actorInfo))];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    _Social.prototype.createSocialActorInfoRequestBody = function (actorInfo) {
        return Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
            "actor": Object.assign(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_4__["metadata"])("SP.Social.SocialActorInfo"), {
                Id: null,
            }, actorInfo),
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("soc.getFollowedSitesUri")
    ], _Social.prototype, "getFollowedSitesUri", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("soc.getFollowedDocumentsUri")
    ], _Social.prototype, "getFollowedDocumentsUri", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("soc.follow")
    ], _Social.prototype, "follow", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("soc.isFollowed")
    ], _Social.prototype, "isFollowed", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("soc.stopFollowing")
    ], _Social.prototype, "stopFollowing", null);
    _Social = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_2__["defaultPath"])("_api/social.following")
    ], _Social);
    return _Social;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

/**
 * Get a new Social instance for the particular Url
 */
var Social = function (baseUrl) { return new _Social(baseUrl); };
var SocialCloneFactory = function (baseUrl, paths) { return new _Social(baseUrl, paths); };
/**
 * Current user's Social instance
 */
var _MySocial = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_MySocial, _super);
    function _MySocial() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    _MySocial.prototype.followed = function (types) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var r;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(MySocialCloneFactory, "followed(types=" + types + ")")()];
                    case 1:
                        r = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["hOP"])(r, "Followed") ? r.Followed.results : r];
                }
            });
        });
    };
    _MySocial.prototype.followedCount = function (types) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var r;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(MySocialCloneFactory, "followedcount(types=" + types + ")")()];
                    case 1:
                        r = _a.sent();
                        return [2 /*return*/, r.FollowedCount || r];
                }
            });
        });
    };
    _MySocial.prototype.followers = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var r;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(MySocialCloneFactory, "followers")()];
                    case 1:
                        r = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["hOP"])(r, "Followers") ? r.Followers.results : r];
                }
            });
        });
    };
    _MySocial.prototype.suggestions = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var r;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(MySocialCloneFactory, "suggestions")()];
                    case 1:
                        r = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["hOP"])(r, "Suggestions") ? r.Suggestions.results : r];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("msoc.followed")
    ], _MySocial.prototype, "followed", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("msoc.followedCount")
    ], _MySocial.prototype, "followedCount", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("msoc.followers")
    ], _MySocial.prototype, "followers", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("msoc.suggestions")
    ], _MySocial.prototype, "suggestions", null);
    _MySocial = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_2__["defaultPath"])("my")
    ], _MySocial);
    return _MySocial;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

/**
 * Invokable factory for IMySocial instances
 */
var MySocial = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_MySocial);
var MySocialCloneFactory = function (baseUrl, path) { return MySocial(baseUrl, path); };
/**
 * Social actor type
 *
 */
var SocialActorType;
(function (SocialActorType) {
    SocialActorType[SocialActorType["User"] = 0] = "User";
    SocialActorType[SocialActorType["Document"] = 1] = "Document";
    SocialActorType[SocialActorType["Site"] = 2] = "Site";
    SocialActorType[SocialActorType["Tag"] = 3] = "Tag";
})(SocialActorType || (SocialActorType = {}));
/**
 * Social actor type
 *
 */
/* tslint:disable:no-bitwise */
var SocialActorTypes;
(function (SocialActorTypes) {
    SocialActorTypes[SocialActorTypes["None"] = 0] = "None";
    SocialActorTypes[SocialActorTypes["User"] = 1] = "User";
    SocialActorTypes[SocialActorTypes["Document"] = 2] = "Document";
    SocialActorTypes[SocialActorTypes["Site"] = 4] = "Site";
    SocialActorTypes[SocialActorTypes["Tag"] = 8] = "Tag";
    /**
     * The set excludes documents and sites that do not have feeds.
     */
    SocialActorTypes[SocialActorTypes["ExcludeContentWithoutFeeds"] = 268435456] = "ExcludeContentWithoutFeeds";
    /**
     * The set includes group sites
     */
    SocialActorTypes[SocialActorTypes["IncludeGroupsSites"] = 536870912] = "IncludeGroupsSites";
    /**
     * The set includes only items created within the last 24 hours
     */
    SocialActorTypes[SocialActorTypes["WithinLast24Hours"] = 1073741824] = "WithinLast24Hours";
})(SocialActorTypes || (SocialActorTypes = {}));
/* tslint:enable */
/**
 * Result from following
 *
 */
var SocialFollowResult;
(function (SocialFollowResult) {
    SocialFollowResult[SocialFollowResult["Ok"] = 0] = "Ok";
    SocialFollowResult[SocialFollowResult["AlreadyFollowing"] = 1] = "AlreadyFollowing";
    SocialFollowResult[SocialFollowResult["LimitReached"] = 2] = "LimitReached";
    SocialFollowResult[SocialFollowResult["InternalError"] = 3] = "InternalError";
})(SocialFollowResult || (SocialFollowResult = {}));
/**
 * Specifies an exception or status code.
 */
var SocialStatusCode;
(function (SocialStatusCode) {
    /**
     * The operation completed successfully
     */
    SocialStatusCode[SocialStatusCode["OK"] = 0] = "OK";
    /**
     * The request is invalid.
     */
    SocialStatusCode[SocialStatusCode["InvalidRequest"] = 1] = "InvalidRequest";
    /**
     *  The current user is not authorized to perform the operation.
     */
    SocialStatusCode[SocialStatusCode["AccessDenied"] = 2] = "AccessDenied";
    /**
     * The target of the operation was not found.
     */
    SocialStatusCode[SocialStatusCode["ItemNotFound"] = 3] = "ItemNotFound";
    /**
     * The operation is invalid for the target's current state.
     */
    SocialStatusCode[SocialStatusCode["InvalidOperation"] = 4] = "InvalidOperation";
    /**
     * The operation completed without modifying the target.
     */
    SocialStatusCode[SocialStatusCode["ItemNotModified"] = 5] = "ItemNotModified";
    /**
     * The operation failed because an internal error occurred.
     */
    SocialStatusCode[SocialStatusCode["InternalError"] = 6] = "InternalError";
    /**
     * The operation failed because the server could not access the distributed cache.
     */
    SocialStatusCode[SocialStatusCode["CacheReadError"] = 7] = "CacheReadError";
    /**
     * The operation succeeded but the server could not update the distributed cache.
     */
    SocialStatusCode[SocialStatusCode["CacheUpdateError"] = 8] = "CacheUpdateError";
    /**
     * No personal site exists for the current user, and no further information is available.
     */
    SocialStatusCode[SocialStatusCode["PersonalSiteNotFound"] = 9] = "PersonalSiteNotFound";
    /**
     * No personal site exists for the current user, and a previous attempt to create one failed.
     */
    SocialStatusCode[SocialStatusCode["FailedToCreatePersonalSite"] = 10] = "FailedToCreatePersonalSite";
    /**
     * No personal site exists for the current user, and a previous attempt to create one was not authorized.
     */
    SocialStatusCode[SocialStatusCode["NotAuthorizedToCreatePersonalSite"] = 11] = "NotAuthorizedToCreatePersonalSite";
    /**
     * No personal site exists for the current user, and no attempt should be made to create one.
     */
    SocialStatusCode[SocialStatusCode["CannotCreatePersonalSite"] = 12] = "CannotCreatePersonalSite";
    /**
     * The operation was rejected because an internal limit had been reached.
     */
    SocialStatusCode[SocialStatusCode["LimitReached"] = 13] = "LimitReached";
    /**
     * The operation failed because an error occurred during the processing of the specified attachment.
     */
    SocialStatusCode[SocialStatusCode["AttachmentError"] = 14] = "AttachmentError";
    /**
     * The operation succeeded with recoverable errors; the returned data is incomplete.
     */
    SocialStatusCode[SocialStatusCode["PartialData"] = 15] = "PartialData";
    /**
     * A required SharePoint feature is not enabled.
     */
    SocialStatusCode[SocialStatusCode["FeatureDisabled"] = 16] = "FeatureDisabled";
    /**
     * The site's storage quota has been exceeded.
     */
    SocialStatusCode[SocialStatusCode["StorageQuotaExceeded"] = 17] = "StorageQuotaExceeded";
    /**
     * The operation failed because the server could not access the database.
     */
    SocialStatusCode[SocialStatusCode["DatabaseError"] = 18] = "DatabaseError";
})(SocialStatusCode || (SocialStatusCode = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "6k7F":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/webs/index.js ***!
  \********************************************/
/*! exports provided: Web, Webs */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./types */ "dVsc");
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Web", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["Web"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Webs", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["Webs"]; });




Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_1__["SPRest"].prototype, "web", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_0__["Web"])(this._baseUrl).configure(this._options);
    },
});
_rest__WEBPACK_IMPORTED_MODULE_1__["SPRest"].prototype.createBatch = function () {
    return this.web.createBatch();
};
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "6k8n":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/fields/list.js ***!
  \*********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "KYlJ");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "fields", _types__WEBPACK_IMPORTED_MODULE_2__["Fields"]);
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "6l60":
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/site-groups/index.js ***!
  \***************************************************/
/*! exports provided: SiteGroup, SiteGroups */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "0rHm");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "UuUm");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteGroup", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["SiteGroup"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteGroups", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["SiteGroups"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "7pU6":
/*!**********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/site.js ***!
  \**********************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sites_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sites/types */ "ilAS");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "1mTr");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_sites_types__WEBPACK_IMPORTED_MODULE_1__["_Site"], "userCustomActions", _types__WEBPACK_IMPORTED_MODULE_2__["UserCustomActions"]);
//# sourceMappingURL=site.js.map

/***/ }),

/***/ "8gyt":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/web.js ***!
  \*********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "AvPO");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./funcs */ "At/s");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../operations */ "UK2s");









/**
 * Shares this web with the supplied users
 * @param loginNames The resolved login names to share
 * @param role The role to share this web
 * @param emailData Optional email data
 */
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.shareWith = function (loginNames, role, emailData) {
    if (role === void 0) { role = _types__WEBPACK_IMPORTED_MODULE_2__["SharingRole"].View; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, web, url;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    web = new _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableInstance"](Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__["extractWebUrl"])(this.toUrl()), "/_api/web/url");
                    return [4 /*yield*/, web.get()];
                case 1:
                    url = _a.sent();
                    dependency();
                    return [2 /*return*/, this.shareObject(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_6__["combine"])(url, "/_layouts/15/aclinv.aspx?forSharing=1&mbypass=1"), loginNames, role, emailData)];
            }
        });
    });
};
/**
 * Provides direct access to the static web.ShareObject method
 *
 * @param url The url to share
 * @param loginNames Resolved loginnames string[] of a single login name string
 * @param roleValue Role value
 * @param emailData Optional email data
 * @param groupId Optional group id
 * @param propagateAcl
 * @param includeAnonymousLinkInEmail
 * @param useSimplifiedRoles
 */
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.shareObject = function (url, loginNames, role, emailData, group, propagateAcl, includeAnonymousLinkInEmail, useSimplifiedRoles) {
    if (propagateAcl === void 0) { propagateAcl = false; }
    if (includeAnonymousLinkInEmail === void 0) { includeAnonymousLinkInEmail = false; }
    if (useSimplifiedRoles === void 0) { useSimplifiedRoles = true; }
    return Object(_funcs__WEBPACK_IMPORTED_MODULE_5__["shareObject"])(this, {
        emailData: emailData,
        group: group,
        includeAnonymousLinkInEmail: includeAnonymousLinkInEmail,
        loginNames: loginNames,
        propagateAcl: propagateAcl,
        role: role,
        url: url,
        useSimplifiedRoles: useSimplifiedRoles,
    });
};
/**
 * Supplies a method to pass any set of arguments to ShareObject
 *
 * @param options The set of options to send to ShareObject
 */
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.shareObjectRaw = function (options) {
    return Object(_funcs__WEBPACK_IMPORTED_MODULE_5__["shareObject"])(this, options, true);
};
/**
 * Supplies a method to pass any set of arguments to ShareObject
 *
 * @param options The set of options to send to ShareObject
 */
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.unshareObject = function (url) {
    return Object(_operations__WEBPACK_IMPORTED_MODULE_8__["spPost"])(Object(_webs_types__WEBPACK_IMPORTED_MODULE_1__["Web"])(this, "unshareObject"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_7__["body"])({ url: url }));
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "8xxP":
/*!***************************************!*\
  !*** ./node_modules/@pnp/sp/batch.js ***!
  \***************************************/
/*! exports provided: SPBatch */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SPBatch", function() { return SPBatch; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_logging__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/logging */ "B3L+");
/* harmony import */ var _sphttpclient__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./sphttpclient */ "DqvY");
/* harmony import */ var _splibconfig__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./splibconfig */ "0p/9");
/* harmony import */ var _utils_toabsoluteurl__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./utils/toabsoluteurl */ "G2Ar");







/**
 * Manages a batch of OData operations
 */
var SPBatch = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(SPBatch, _super);
    function SPBatch(baseUrl) {
        var _this = _super.call(this) || this;
        _this.baseUrl = baseUrl;
        return _this;
    }
    /**
     * Parses the response from a batch request into an array of Response instances
     *
     * @param body Text body of the response from the batch request
     */
    SPBatch.ParseResponse = function (body) {
        var responses = [];
        var header = "--batchresponse_";
        // Ex. "HTTP/1.1 500 Internal Server Error"
        var statusRegExp = new RegExp("^HTTP/[0-9.]+ +([0-9]+) +(.*)", "i");
        var lines = body.split("\n");
        var state = "batch";
        var status;
        var statusText;
        for (var i = 0; i < lines.length; ++i) {
            var line = lines[i];
            switch (state) {
                case "batch":
                    if (line.substr(0, header.length) === header) {
                        state = "batchHeaders";
                    }
                    else {
                        if (line.trim() !== "") {
                            throw Error("Invalid response, line " + i);
                        }
                    }
                    break;
                case "batchHeaders":
                    if (line.trim() === "") {
                        state = "status";
                    }
                    break;
                case "status":
                    var parts = statusRegExp.exec(line);
                    if (parts.length !== 3) {
                        throw Error("Invalid status, line " + i);
                    }
                    status = parseInt(parts[1], 10);
                    statusText = parts[2];
                    state = "statusHeaders";
                    break;
                case "statusHeaders":
                    if (line.trim() === "") {
                        state = "body";
                    }
                    break;
                case "body":
                    responses.push((status === 204) ? new Response() : new Response(line, { status: status, statusText: statusText }));
                    state = "batch";
                    break;
            }
        }
        if (state !== "status") {
            throw Error("Unexpected end of input");
        }
        return responses;
    };
    SPBatch.prototype.executeImpl = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var client, absoluteRequestUrl, batchBody, currentChangeSetId, i, reqInfo, headers, url, method, castHeaders, batchOptions, fetchResponse, text, responses;
            var _this = this;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _pnp_logging__WEBPACK_IMPORTED_MODULE_3__["Logger"].write("[" + this.batchId + "] (" + (new Date()).getTime() + ") Executing batch with " + this.requests.length + " requests.", 1 /* Info */);
                        // if we don't have any requests, don't bother sending anything
                        // this could be due to caching further upstream, or just an empty batch
                        if (this.requests.length < 1) {
                            _pnp_logging__WEBPACK_IMPORTED_MODULE_3__["Logger"].write("Resolving empty batch.", 1 /* Info */);
                            return [2 /*return*/];
                        }
                        client = new _sphttpclient__WEBPACK_IMPORTED_MODULE_4__["SPHttpClient"]();
                        return [4 /*yield*/, Object(_utils_toabsoluteurl__WEBPACK_IMPORTED_MODULE_6__["toAbsoluteUrl"])(this.baseUrl)];
                    case 1:
                        absoluteRequestUrl = _a.sent();
                        batchBody = [];
                        currentChangeSetId = "";
                        for (i = 0; i < this.requests.length; i++) {
                            reqInfo = this.requests[i];
                            if (reqInfo.method === "GET") {
                                if (currentChangeSetId.length > 0) {
                                    // end an existing change set
                                    batchBody.push("--changeset_" + currentChangeSetId + "--\n\n");
                                    currentChangeSetId = "";
                                }
                                batchBody.push("--batch_" + this.batchId + "\n");
                            }
                            else {
                                if (currentChangeSetId.length < 1) {
                                    // start new change set
                                    currentChangeSetId = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["getGUID"])();
                                    batchBody.push("--batch_" + this.batchId + "\n");
                                    batchBody.push("Content-Type: multipart/mixed; boundary=\"changeset_" + currentChangeSetId + "\"\n\n");
                                }
                                batchBody.push("--changeset_" + currentChangeSetId + "\n");
                            }
                            // common batch part prefix
                            batchBody.push("Content-Type: application/http\n");
                            batchBody.push("Content-Transfer-Encoding: binary\n\n");
                            headers = new Headers();
                            url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["isUrlAbsolute"])(reqInfo.url) ? reqInfo.url : Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["combine"])(absoluteRequestUrl, reqInfo.url);
                            _pnp_logging__WEBPACK_IMPORTED_MODULE_3__["Logger"].write("[" + this.batchId + "] (" + (new Date()).getTime() + ") Adding request " + reqInfo.method + " " + url + " to batch.", 0 /* Verbose */);
                            if (reqInfo.method !== "GET") {
                                method = reqInfo.method;
                                castHeaders = reqInfo.options.headers;
                                if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(reqInfo, "options") && Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(reqInfo.options, "headers") && castHeaders["X-HTTP-Method"] !== undefined) {
                                    method = castHeaders["X-HTTP-Method"];
                                    delete castHeaders["X-HTTP-Method"];
                                }
                                batchBody.push(method + " " + url + " HTTP/1.1\n");
                                headers.set("Content-Type", "application/json;odata=verbose;charset=utf-8");
                            }
                            else {
                                batchBody.push(reqInfo.method + " " + url + " HTTP/1.1\n");
                            }
                            // merge global config headers
                            Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["mergeHeaders"])(headers, _splibconfig__WEBPACK_IMPORTED_MODULE_5__["SPRuntimeConfig"].headers);
                            // merge per-request headers
                            if (reqInfo.options) {
                                Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["mergeHeaders"])(headers, reqInfo.options.headers);
                            }
                            // lastly we apply any default headers we need that may not exist
                            if (!headers.has("Accept")) {
                                headers.append("Accept", "application/json");
                            }
                            if (!headers.has("Content-Type")) {
                                headers.append("Content-Type", "application/json;odata=verbose;charset=utf-8");
                            }
                            if (!headers.has("X-ClientService-ClientTag")) {
                                headers.append("X-ClientService-ClientTag", "PnPCoreJS:@pnp-2.0.5:batch");
                            }
                            // write headers into batch body
                            headers.forEach(function (value, name) {
                                batchBody.push(name + ": " + value + "\n");
                            });
                            batchBody.push("\n");
                            if (reqInfo.options.body) {
                                batchBody.push(reqInfo.options.body + "\n\n");
                            }
                        }
                        if (currentChangeSetId.length > 0) {
                            // Close the changeset
                            batchBody.push("--changeset_" + currentChangeSetId + "--\n\n");
                            currentChangeSetId = "";
                        }
                        batchBody.push("--batch_" + this.batchId + "--\n");
                        batchOptions = {
                            "body": batchBody.join(""),
                            "headers": {
                                "Content-Type": "multipart/mixed; boundary=batch_" + this.batchId,
                            },
                            "method": "POST",
                        };
                        _pnp_logging__WEBPACK_IMPORTED_MODULE_3__["Logger"].write("[" + this.batchId + "] (" + (new Date()).getTime() + ") Sending batch request.", 1 /* Info */);
                        return [4 /*yield*/, client.fetch(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["combine"])(absoluteRequestUrl, "/_api/$batch"), batchOptions)];
                    case 2:
                        fetchResponse = _a.sent();
                        return [4 /*yield*/, fetchResponse.text()];
                    case 3:
                        text = _a.sent();
                        responses = SPBatch.ParseResponse(text);
                        if (responses.length !== this.requests.length) {
                            throw Error("Could not properly parse responses to match requests in batch.");
                        }
                        _pnp_logging__WEBPACK_IMPORTED_MODULE_3__["Logger"].write("[" + this.batchId + "] (" + (new Date()).getTime() + ") Resolving batched requests.", 1 /* Info */);
                        // this structure ensures that we resolve the batched requests in the order we expect
                        // using async this is not guaranteed depending on the requests
                        return [2 /*return*/, responses.reduce(function (p, response, index) { return p.then(function (_) { return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(_this, void 0, void 0, function () {
                                var request, _a, _b, e_1;
                                return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_c) {
                                    switch (_c.label) {
                                        case 0:
                                            request = this.requests[index];
                                            _pnp_logging__WEBPACK_IMPORTED_MODULE_3__["Logger"].write("[" + request.id + "] (" + (new Date()).getTime() + ") Resolving request in batch " + this.batchId + ".", 1 /* Info */);
                                            _c.label = 1;
                                        case 1:
                                            _c.trys.push([1, 3, , 4]);
                                            _b = (_a = request).resolve;
                                            return [4 /*yield*/, request.parser.parse(response)];
                                        case 2:
                                            _b.apply(_a, [_c.sent()]);
                                            return [3 /*break*/, 4];
                                        case 3:
                                            e_1 = _c.sent();
                                            request.reject(e_1);
                                            return [3 /*break*/, 4];
                                        case 4: return [2 /*return*/];
                                    }
                                });
                            }); }); }, Promise.resolve(void (0)))];
                }
            });
        });
    };
    return SPBatch;
}(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["Batch"]));

//# sourceMappingURL=batch.js.map

/***/ }),

/***/ "97IE":
/*!******************************************!*\
  !*** ./node_modules/@pnp/common/util.js ***!
  \******************************************/
/*! exports provided: getCtxCallback, dateAdd, combine, getRandomString, getGUID, isFunc, objectDefinedNotNull, isArray, assign, isUrlAbsolute, stringIsNullOrEmpty, sanitizeGuid, jsS, hOP, getHashCode */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCtxCallback", function() { return getCtxCallback; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "dateAdd", function() { return dateAdd; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "combine", function() { return combine; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getRandomString", function() { return getRandomString; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getGUID", function() { return getGUID; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isFunc", function() { return isFunc; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "objectDefinedNotNull", function() { return objectDefinedNotNull; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isArray", function() { return isArray; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "assign", function() { return assign; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "isUrlAbsolute", function() { return isUrlAbsolute; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "stringIsNullOrEmpty", function() { return stringIsNullOrEmpty; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sanitizeGuid", function() { return sanitizeGuid; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "jsS", function() { return jsS; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "hOP", function() { return hOP; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getHashCode", function() { return getHashCode; });
/**
 * Gets a callback function which will maintain context across async calls.
 * Allows for the calling pattern getCtxCallback(thisobj, method, methodarg1, methodarg2, ...)
 *
 * @param context The object that will be the 'this' value in the callback
 * @param method The method to which we will apply the context and parameters
 * @param params Optional, additional arguments to supply to the wrapped method when it is invoked
 */
function getCtxCallback(context, method) {
    var params = [];
    for (var _i = 2; _i < arguments.length; _i++) {
        params[_i - 2] = arguments[_i];
    }
    return function () {
        method.apply(context, params);
    };
}
/**
 * Adds a value to a date
 *
 * @param date The date to which we will add units, done in local time
 * @param interval The name of the interval to add, one of: ['year', 'quarter', 'month', 'week', 'day', 'hour', 'minute', 'second']
 * @param units The amount to add to date of the given interval
 *
 * http://stackoverflow.com/questions/1197928/how-to-add-30-minutes-to-a-javascript-date-object
 */
function dateAdd(date, interval, units) {
    var ret = new Date(date.toString()); // don't change original date
    switch (interval.toLowerCase()) {
        case "year":
            ret.setFullYear(ret.getFullYear() + units);
            break;
        case "quarter":
            ret.setMonth(ret.getMonth() + 3 * units);
            break;
        case "month":
            ret.setMonth(ret.getMonth() + units);
            break;
        case "week":
            ret.setDate(ret.getDate() + 7 * units);
            break;
        case "day":
            ret.setDate(ret.getDate() + units);
            break;
        case "hour":
            ret.setTime(ret.getTime() + units * 3600000);
            break;
        case "minute":
            ret.setTime(ret.getTime() + units * 60000);
            break;
        case "second":
            ret.setTime(ret.getTime() + units * 1000);
            break;
        default:
            ret = undefined;
            break;
    }
    return ret;
}
/**
 * Combines an arbitrary set of paths ensuring and normalizes the slashes
 *
 * @param paths 0 to n path parts to combine
 */
function combine() {
    var paths = [];
    for (var _i = 0; _i < arguments.length; _i++) {
        paths[_i] = arguments[_i];
    }
    return paths
        .filter(function (path) { return !stringIsNullOrEmpty(path); })
        .map(function (path) { return path.replace(/^[\\|\/]/, "").replace(/[\\|\/]$/, ""); })
        .join("/")
        .replace(/\\/g, "/");
}
/**
 * Gets a random string of chars length
 *
 * https://stackoverflow.com/questions/1349404/generate-random-string-characters-in-javascript
 *
 * @param chars The length of the random string to generate
 */
function getRandomString(chars) {
    var text = new Array(chars);
    for (var i = 0; i < chars; i++) {
        text[i] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789".charAt(Math.floor(Math.random() * 62));
    }
    return text.join("");
}
/**
 * Gets a random GUID value
 *
 * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
 */
/* tslint:disable no-bitwise */
function getGUID() {
    var d = Date.now();
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, function (c) {
        var r = (d + Math.random() * 16) % 16 | 0;
        d = Math.floor(d / 16);
        return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
    });
}
/* tslint:enable */
/**
 * Determines if a given value is a function
 *
 * @param cf The thing to test for functionness
 */
function isFunc(f) {
    return typeof f === "function";
}
/**
 * Determines if an object is both defined and not null
 * @param obj Object to test
 */
function objectDefinedNotNull(obj) {
    return typeof obj !== "undefined" && obj !== null;
}
/**
 * @returns whether the provided parameter is a JavaScript Array or not.
*/
function isArray(array) {
    return Array.isArray ? Array.isArray(array) : array && typeof array.length === "number" && array.constructor === Array;
}
/**
 * Provides functionality to extend the given object by doing a shallow copy
 *
 * @param target The object to which properties will be copied
 * @param source The source object from which properties will be copied
 * @param noOverwrite If true existing properties on the target are not overwritten from the source
 * @param filter If provided allows additional filtering on what properties are copied (propName: string) => boolean
 *
 */
function assign(target, source, noOverwrite, filter) {
    if (noOverwrite === void 0) { noOverwrite = false; }
    if (filter === void 0) { filter = function () { return true; }; }
    if (!objectDefinedNotNull(source)) {
        return target;
    }
    // ensure we don't overwrite things we don't want overwritten
    var check = noOverwrite ? function (o, i) { return !(i in o); } : function () { return true; };
    // final filter we will use
    var f = function (v) { return check(target, v) && filter(v); };
    return Object.getOwnPropertyNames(source)
        .filter(f)
        .reduce(function (t, v) {
        t[v] = source[v];
        return t;
    }, target);
}
/**
 * Determines if a given url is absolute
 *
 * @param url The url to check to see if it is absolute
 */
function isUrlAbsolute(url) {
    return /^https?:\/\/|^\/\//i.test(url);
}
/**
 * Determines if a string is null or empty or undefined
 *
 * @param s The string to test
 */
function stringIsNullOrEmpty(s) {
    return s === undefined || s === null || s.length < 1;
}
/**
 * Ensures guid values are represented consistently as "ea123463-137d-4ae3-89b8-cf3fc578ca05"
 *
 * @param guid The candidate guid
 */
function sanitizeGuid(guid) {
    if (stringIsNullOrEmpty(guid)) {
        return guid;
    }
    var matches = /([0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12})/i.exec(guid);
    return matches === null ? guid : matches[1];
}
/**
 * Shorthand for JSON.stringify
 *
 * @param o Any type of object
 */
function jsS(o) {
    return JSON.stringify(o);
}
/**
 * Shorthand for Object.hasOwnProperty
 *
 * @param o Object to check for
 * @param p Name of the property
 */
function hOP(o, p) {
    return Object.hasOwnProperty.call(o, p);
}
/**
 * Generates a ~unique hash code
 *
 * From: https://stackoverflow.com/questions/6122571/simple-non-secure-hash-function-for-javascript
 */
// tslint:disable:no-bitwise
function getHashCode(s) {
    var hash = 0;
    if (s.length === 0) {
        return hash;
    }
    for (var i = 0; i < s.length; i++) {
        var chr = s.charCodeAt(i);
        hash = ((hash << 5) - hash) + chr;
        hash |= 0; // Convert to 32bit integer
    }
    return hash;
}
// tslint:enable:no-bitwise
//# sourceMappingURL=util.js.map

/***/ }),

/***/ "9Ar0":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/webparts/index.js ***!
  \************************************************/
/*! exports provided: WebPartsPersonalizationScope, WebPartDefinitions, WebPartDefinition, LimitedWebPartManager */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _file__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./file */ "yGnx");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "pXI8");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "WebPartsPersonalizationScope", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["WebPartsPersonalizationScope"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "WebPartDefinitions", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["WebPartDefinitions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "WebPartDefinition", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["WebPartDefinition"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "LimitedWebPartManager", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["LimitedWebPartManager"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "9TEK":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/common/storage.js ***!
  \*********************************************/
/*! exports provided: PnPClientStorageWrapper, PnPClientStorage */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PnPClientStorageWrapper", function() { return PnPClientStorageWrapper; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PnPClientStorage", function() { return PnPClientStorage; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "lzYu");
/* harmony import */ var _util__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./util */ "97IE");
/* harmony import */ var _libconfig__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./libconfig */ "hyGn");



/**
 * A wrapper class to provide a consistent interface to browser based storage
 *
 */
var PnPClientStorageWrapper = /** @class */ (function () {
    /**
     * Creates a new instance of the PnPClientStorageWrapper class
     *
     * @constructor
     */
    function PnPClientStorageWrapper(store, defaultTimeoutMinutes) {
        if (defaultTimeoutMinutes === void 0) { defaultTimeoutMinutes = -1; }
        this.store = store;
        this.defaultTimeoutMinutes = defaultTimeoutMinutes;
        this.enabled = this.test();
        // if the cache timeout is enabled call the handler
        // this will clear any expired items and set the timeout function
        if (_libconfig__WEBPACK_IMPORTED_MODULE_2__["RuntimeConfig"].enableCacheExpiration) {
            this.cacheExpirationHandler();
        }
    }
    PnPClientStorageWrapper.bind = function (store) {
        return new PnPClientStorageWrapper(typeof (store) === "undefined" ? new MemoryStorage() : store);
    };
    /**
     * Get a value from storage, or null if that value does not exist
     *
     * @param key The key whose value we want to retrieve
     */
    PnPClientStorageWrapper.prototype.get = function (key) {
        if (!this.enabled) {
            return null;
        }
        var o = this.store.getItem(key);
        if (!Object(_util__WEBPACK_IMPORTED_MODULE_1__["objectDefinedNotNull"])(o)) {
            return null;
        }
        var persistable = JSON.parse(o);
        if (new Date(persistable.expiration) <= new Date()) {
            this.delete(key);
            return null;
        }
        else {
            return persistable.value;
        }
    };
    /**
     * Adds a value to the underlying storage
     *
     * @param key The key to use when storing the provided value
     * @param o The value to store
     * @param expire Optional, if provided the expiration of the item, otherwise the default is used
     */
    PnPClientStorageWrapper.prototype.put = function (key, o, expire) {
        if (this.enabled) {
            this.store.setItem(key, this.createPersistable(o, expire));
        }
    };
    /**
     * Deletes a value from the underlying storage
     *
     * @param key The key of the pair we want to remove from storage
     */
    PnPClientStorageWrapper.prototype.delete = function (key) {
        if (this.enabled) {
            this.store.removeItem(key);
        }
    };
    /**
     * Gets an item from the underlying storage, or adds it if it does not exist using the supplied getter function
     *
     * @param key The key to use when storing the provided value
     * @param getter A function which will upon execution provide the desired value
     * @param expire Optional, if provided the expiration of the item, otherwise the default is used
     */
    PnPClientStorageWrapper.prototype.getOrPut = function (key, getter, expire) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var o;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.enabled) {
                            return [2 /*return*/, getter()];
                        }
                        o = this.get(key);
                        if (!(o === null)) return [3 /*break*/, 2];
                        return [4 /*yield*/, getter()];
                    case 1:
                        o = _a.sent();
                        this.put(key, o, expire);
                        _a.label = 2;
                    case 2: return [2 /*return*/, o];
                }
            });
        });
    };
    /**
     * Deletes any expired items placed in the store by the pnp library, leaves other items untouched
     */
    PnPClientStorageWrapper.prototype.deleteExpired = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var i, key;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.enabled) {
                            return [2 /*return*/];
                        }
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!(i < this.store.length)) return [3 /*break*/, 4];
                        key = this.store.key(i);
                        if (!(key !== null)) return [3 /*break*/, 3];
                        if (!/["|']?pnp["|']? ?: ?1/i.test(this.store.getItem(key))) return [3 /*break*/, 3];
                        // get those items as get will delete from cache if they are expired
                        return [4 /*yield*/, this.get(key)];
                    case 2:
                        // get those items as get will delete from cache if they are expired
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Used to determine if the wrapped storage is available currently
     */
    PnPClientStorageWrapper.prototype.test = function () {
        var str = "t";
        try {
            this.store.setItem(str, str);
            this.store.removeItem(str);
            return true;
        }
        catch (e) {
            return false;
        }
    };
    /**
     * Creates the persistable to store
     */
    PnPClientStorageWrapper.prototype.createPersistable = function (o, expire) {
        if (expire === undefined) {
            // ensure we are by default inline with the global library setting
            var defaultTimeout = _libconfig__WEBPACK_IMPORTED_MODULE_2__["RuntimeConfig"].defaultCachingTimeoutSeconds;
            if (this.defaultTimeoutMinutes > 0) {
                defaultTimeout = this.defaultTimeoutMinutes * 60;
            }
            expire = Object(_util__WEBPACK_IMPORTED_MODULE_1__["dateAdd"])(new Date(), "second", defaultTimeout);
        }
        return Object(_util__WEBPACK_IMPORTED_MODULE_1__["jsS"])({ pnp: 1, expiration: expire, value: o });
    };
    /**
     * Deletes expired items added by this library in this.store and sets a timeout to call itself
     */
    PnPClientStorageWrapper.prototype.cacheExpirationHandler = function () {
        var _this = this;
        if (!this.enabled) {
            return;
        }
        this.deleteExpired().then(function (_) {
            // call ourself in the future
            setTimeout(Object(_util__WEBPACK_IMPORTED_MODULE_1__["getCtxCallback"])(_this, _this.cacheExpirationHandler), _libconfig__WEBPACK_IMPORTED_MODULE_2__["RuntimeConfig"].cacheExpirationIntervalMilliseconds);
        }).catch(console.error);
    };
    return PnPClientStorageWrapper;
}());

/**
 * A thin implementation of in-memory storage for use in nodejs
 */
var MemoryStorage = /** @class */ (function () {
    function MemoryStorage(_store) {
        if (_store === void 0) { _store = new Map(); }
        this._store = _store;
    }
    Object.defineProperty(MemoryStorage.prototype, "length", {
        get: function () {
            return this._store.size;
        },
        enumerable: true,
        configurable: true
    });
    MemoryStorage.prototype.clear = function () {
        this._store.clear();
    };
    MemoryStorage.prototype.getItem = function (key) {
        return this._store.get(key);
    };
    MemoryStorage.prototype.key = function (index) {
        return Array.from(this._store)[index][0];
    };
    MemoryStorage.prototype.removeItem = function (key) {
        this._store.delete(key);
    };
    MemoryStorage.prototype.setItem = function (key, data) {
        this._store.set(key, data);
    };
    return MemoryStorage;
}());
/**
 * A class that will establish wrappers for both local and session storage
 */
var PnPClientStorage = /** @class */ (function () {
    /**
     * Creates a new instance of the PnPClientStorage class
     *
     * @constructor
     */
    function PnPClientStorage(_local, _session) {
        if (_local === void 0) { _local = null; }
        if (_session === void 0) { _session = null; }
        this._local = _local;
        this._session = _session;
    }
    Object.defineProperty(PnPClientStorage.prototype, "local", {
        /**
         * Provides access to the local storage of the browser
         */
        get: function () {
            if (this._local === null) {
                this._local = new PnPClientStorageWrapper(typeof (localStorage) === "undefined" ? new MemoryStorage() : localStorage);
            }
            return this._local;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(PnPClientStorage.prototype, "session", {
        /**
         * Provides access to the session storage of the browser
         */
        get: function () {
            if (this._session === null) {
                this._session = new PnPClientStorageWrapper(typeof (sessionStorage) === "undefined" ? new MemoryStorage() : sessionStorage);
            }
            return this._session;
        },
        enumerable: true,
        configurable: true
    });
    return PnPClientStorage;
}());

//# sourceMappingURL=storage.js.map

/***/ }),

/***/ "9l4K":
/*!**************************************!*\
  !*** ./node_modules/@pnp/sp/rest.js ***!
  \**************************************/
/*! exports provided: SPRest, sp */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SPRest", function() { return SPRest; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sp", function() { return sp; });
/* harmony import */ var _splibconfig__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./splibconfig */ "0p/9");

/**
 * Root of the SharePoint REST module
 */
var SPRest = /** @class */ (function () {
    /**
     * Creates a new instance of the SPRest class
     *
     * @param options Additional options
     * @param baseUrl A string that should form the base part of the url
     */
    function SPRest(_options, _baseUrl) {
        if (_options === void 0) { _options = {}; }
        if (_baseUrl === void 0) { _baseUrl = ""; }
        this._options = _options;
        this._baseUrl = _baseUrl;
    }
    /**
     * Configures instance with additional options and baseUrl.
     * Provided configuration used by other objects in a chain
     *
     * @param options Additional options
     * @param baseUrl A string that should form the base part of the url
     */
    SPRest.prototype.configure = function (options, baseUrl) {
        if (baseUrl === void 0) { baseUrl = ""; }
        return new SPRest(options, baseUrl);
    };
    /**
     * Global SharePoint configuration options
     *
     * @param config The SharePoint configuration to apply
     */
    SPRest.prototype.setup = function (config) {
        if (config.pageContext) {
            Object(_splibconfig__WEBPACK_IMPORTED_MODULE_0__["setup"])({
                spfxContext: config,
            });
        }
        else {
            Object(_splibconfig__WEBPACK_IMPORTED_MODULE_0__["setup"])(config);
        }
    };
    return SPRest;
}());

var sp = new SPRest();
//# sourceMappingURL=rest.js.map

/***/ }),

/***/ "AfLD":
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/sputilities/types.js ***!
  \***************************************************/
/*! exports provided: _Utilities, Utilities */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Utilities", function() { return _Utilities; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Utilities", function() { return Utilities; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _files_types__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../files/types */ "qfcF");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _utils_objectToSPKeyValueCollection__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../utils/objectToSPKeyValueCollection */ "EVeO");











var _Utilities = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Utilities, _super);
    function _Utilities(baseUrl, methodName) {
        var _this = this;
        var url = typeof baseUrl === "string" ? baseUrl : baseUrl.toUrl();
        _this = _super.call(this, Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_7__["extractWebUrl"])(url), "_api/SP.Utilities.Utility." + methodName) || this;
        return _this;
    }
    _Utilities.prototype.excute = function (props) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_8__["spPost"])(this, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_3__["body"])(props));
    };
    _Utilities.prototype.sendEmail = function (props) {
        var params = {
            properties: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.Utilities.EmailProperties"), {
                Body: props.Body,
                From: props.From,
                Subject: props.Subject,
            }),
        };
        if (props.To && props.To.length > 0) {
            params.properties = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(params.properties, {
                To: { results: props.To },
            });
        }
        if (props.CC && props.CC.length > 0) {
            params.properties = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(params.properties, {
                CC: { results: props.CC },
            });
        }
        if (props.BCC && props.BCC.length > 0) {
            params.properties = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(params.properties, {
                BCC: { results: props.BCC },
            });
        }
        if (props.AdditionalHeaders) {
            params.properties = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(params.properties, {
                AdditionalHeaders: Object(_utils_objectToSPKeyValueCollection__WEBPACK_IMPORTED_MODULE_10__["objectToSPKeyValueCollection"])(props.AdditionalHeaders),
            });
        }
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(this.clone(UtilitiesCloneFactory, "SendEmail", true), "u.sendEmail").excute(params);
    };
    _Utilities.prototype.getCurrentUserEmailAddresses = function () {
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(this.clone(UtilitiesCloneFactory, "GetCurrentUserEmailAddresses", true), "u.getCurrentUserEmailAddresses").excute({});
    };
    _Utilities.prototype.resolvePrincipal = function (input, scopes, sources, inputIsEmailOnly, addToUserInfoList, matchUserInfoList) {
        if (matchUserInfoList === void 0) { matchUserInfoList = false; }
        var params = {
            addToUserInfoList: addToUserInfoList,
            input: input,
            inputIsEmailOnly: inputIsEmailOnly,
            matchUserInfoList: matchUserInfoList,
            scopes: scopes,
            sources: sources,
        };
        var clone = this.clone(UtilitiesCloneFactory, "ResolvePrincipalInCurrentContext", true);
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(clone, "u.ResolvePrincipalInCurrentContext").excute(params);
    };
    _Utilities.prototype.searchPrincipals = function (input, scopes, sources, groupName, maxCount) {
        var params = {
            groupName: groupName,
            input: input,
            maxCount: maxCount,
            scopes: scopes,
            sources: sources,
        };
        var clone = this.clone(UtilitiesCloneFactory, "SearchPrincipalsUsingContextWeb", true);
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(clone, "u.SearchPrincipalsUsingContextWeb").excute(params);
    };
    _Utilities.prototype.createEmailBodyForInvitation = function (pageAddress) {
        var params = {
            pageAddress: pageAddress,
        };
        var clone = this.clone(UtilitiesCloneFactory, "CreateEmailBodyForInvitation", true);
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(clone, "u.CreateEmailBodyForInvitation").excute(params);
    };
    _Utilities.prototype.expandGroupsToPrincipals = function (inputs, maxCount) {
        if (maxCount === void 0) { maxCount = 30; }
        var params = {
            inputs: inputs,
            maxCount: maxCount,
        };
        var clone = this.clone(UtilitiesCloneFactory, "ExpandGroupsToPrincipals", true);
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(clone, "u.ExpandGroupsToPrincipals").excute(params);
    };
    _Utilities.prototype.createWikiPage = function (info) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var clone, newPage;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        clone = this.clone(UtilitiesCloneFactory, "CreateWikiPageInContextWeb", true);
                        return [4 /*yield*/, _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(clone, "u.CreateWikiPageInContextWeb").excute({ parameters: info })];
                    case 1:
                        newPage = _a.sent();
                        return [2 /*return*/, {
                                data: newPage,
                                file: Object(_files_types__WEBPACK_IMPORTED_MODULE_6__["File"])(Object(_odata__WEBPACK_IMPORTED_MODULE_4__["odataUrlFrom"])(newPage)),
                            }];
                }
            });
        });
    };
    return _Utilities;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryable"]));

var Utilities = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Utilities);
var UtilitiesCloneFactory = function (baseUrl, path) { return Utilities(baseUrl, path); };
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "Al5K":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/subscriptions/index.js ***!
  \*****************************************************/
/*! exports provided: Subscription, Subscriptions */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list */ "dRNj");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "N3o1");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Subscription", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Subscription"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Subscriptions", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Subscriptions"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "Ar68":
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/attachments/types.js ***!
  \***************************************************/
/*! exports provided: _Attachments, Attachments, _Attachment, Attachment */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Attachments", function() { return _Attachments; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Attachments", function() { return Attachments; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Attachment", function() { return _Attachment; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Attachment", function() { return Attachment; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../telemetry */ "h3gW");






var _Attachments = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Attachments, _super);
    function _Attachments() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
    * Gets a Attachment File by filename
    *
    * @param name The name of the file, including extension.
    */
    _Attachments.prototype.getByName = function (name) {
        var f = _telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"].configure(Attachment(this), "ats.getByName");
        f.concat("('" + name + "')");
        return f;
    };
    /**
     * Adds a new attachment to the collection. Not supported for batching.
     *
     * @param name The name of the file, including extension.
     * @param content The Base64 file content.
     */
    _Attachments.prototype.add = function (name, content) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var response;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this.clone(Attachments, "add(FileName='" + name + "')", false), { body: content })];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, {
                                data: response,
                                file: this.getByName(name),
                            }];
                }
            });
        });
    };
    /**
     * Adds multiple new attachment to the collection. Not supported for batching.
     *
     * @param files The collection of files to add
     */
    _Attachments.prototype.addMultiple = function (files) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var i;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!(i < files.length)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.add(files[i].name, files[i].content)];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Delete multiple attachments from the collection. Not supported for batching.
     *
     * @param files The collection of files to delete
     */
    _Attachments.prototype.deleteMultiple = function () {
        var files = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            files[_i] = arguments[_i];
        }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var i;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!(i < files.length)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.getByName(files[i]).delete()];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Delete multiple attachments from the collection and send to recycle bin. Not supported for batching.
     *
     * @param files The collection of files to be deleted and sent to recycle bin
     */
    _Attachments.prototype.recycleMultiple = function () {
        var files = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            files[_i] = arguments[_i];
        }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var i;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!(i < files.length)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.getByName(files[i]).recycle()];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ats.add")
    ], _Attachments.prototype, "add", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ats.addMultiple")
    ], _Attachments.prototype, "addMultiple", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ats.deleteMultiple")
    ], _Attachments.prototype, "deleteMultiple", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ats.recycleMultiple")
    ], _Attachments.prototype, "recycleMultiple", null);
    _Attachments = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_1__["defaultPath"])("AttachmentFiles")
    ], _Attachments);
    return _Attachments;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableCollection"]));

var Attachments = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_Attachments);
var _Attachment = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Attachment, _super);
    function _Attachment() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["deleteableWithETag"])("at");
        return _this;
    }
    /**
     * Gets the contents of the file as text
     *
     */
    _Attachment.prototype.getText = function () {
        return this.getParsed(new _pnp_odata__WEBPACK_IMPORTED_MODULE_4__["TextParser"]());
    };
    /**
     * Gets the contents of the file as a blob, does not work in Node.js
     *
     */
    _Attachment.prototype.getBlob = function () {
        return this.getParsed(new _pnp_odata__WEBPACK_IMPORTED_MODULE_4__["BlobParser"]());
    };
    /**
     * Gets the contents of a file as an ArrayBuffer, works in Node.js
     */
    _Attachment.prototype.getBuffer = function () {
        return this.getParsed(new _pnp_odata__WEBPACK_IMPORTED_MODULE_4__["BufferParser"]());
    };
    /**
     * Gets the contents of a file as an ArrayBuffer, works in Node.js
     */
    _Attachment.prototype.getJSON = function () {
        return this.getParsed(new _pnp_odata__WEBPACK_IMPORTED_MODULE_4__["JSONParser"]());
    };
    /**
     * Sets the content of a file. Not supported for batching
     *
     * @param content The value to set for the file contents
     */
    _Attachment.prototype.setContent = function (content) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this.clone(Attachment, "$value", false), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["headers"])({ "X-HTTP-Method": "PUT" }, {
                            body: content,
                        }))];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, Attachment(this)];
                }
            });
        });
    };
    /**
     * Delete this attachment file and send it to recycle bin
     *
     * @param eTag Value used in the IF-Match header, by default "*"
     */
    _Attachment.prototype.recycle = function (eTag) {
        if (eTag === void 0) { eTag = "*"; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this.clone(Attachment, "recycleObject"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["headers"])({
            "IF-Match": eTag,
            "X-HTTP-Method": "DELETE",
        }));
    };
    _Attachment.prototype.getParsed = function (parser) {
        return this.clone(Attachment, "$value", false).usingParser(parser)();
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("at.getText")
    ], _Attachment.prototype, "getText", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("at.getBlob")
    ], _Attachment.prototype, "getBlob", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("at.getBuffer")
    ], _Attachment.prototype, "getBuffer", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("at.getJSON")
    ], _Attachment.prototype, "getJSON", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("at.setContent")
    ], _Attachment.prototype, "setContent", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("at.recycle")
    ], _Attachment.prototype, "recycle", null);
    return _Attachment;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableInstance"]));

var Attachment = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_Attachment);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "ArVI":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/appcatalog/types.js ***!
  \**************************************************/
/*! exports provided: _AppCatalog, AppCatalog, _App, App */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_AppCatalog", function() { return _AppCatalog; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AppCatalog", function() { return AppCatalog; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_App", function() { return _App; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "App", function() { return App; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _files_types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../files/types */ "qfcF");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _webs__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../webs */ "6k7F");
/* harmony import */ var _items__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../items */ "lYrR");









var _AppCatalog = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_AppCatalog, _super);
    function _AppCatalog(baseUrl, path) {
        if (path === void 0) { path = "_api/web/tenantappcatalog/AvailableApps"; }
        return _super.call(this, Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__["extractWebUrl"])(typeof baseUrl === "string" ? baseUrl : baseUrl.toUrl()), path) || this;
    }
    /**
     * Get details of specific app from the app catalog
     * @param id - Specify the guid of the app
     */
    _AppCatalog.prototype.getAppById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(App(this, "getById('" + id + "')"), "ac.getAppById");
    };
    /**
     * Synchronize a solution to the Microsoft Teams App Catalog
     * @param id - Specify the guid of the app
     * @param useSharePointItemId (optional) - By default this REST call requires the SP Item id of the app, not the app id.
     *                            PnPjs will try to fetch the item id by default, you can still use this parameter to pass your own item id in the first parameter
     */
    _AppCatalog.prototype.syncSolutionToTeams = function (id, useSharePointItemId) {
        if (useSharePointItemId === void 0) { useSharePointItemId = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var appId, webUrl, web, listId, listItems, poster;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        appId = null;
                        webUrl = Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__["extractWebUrl"])(this.toUrl());
                        if (!useSharePointItemId) return [3 /*break*/, 1];
                        appId = id;
                        return [3 /*break*/, 4];
                    case 1:
                        web = Object(_webs__WEBPACK_IMPORTED_MODULE_7__["Web"])(webUrl);
                        return [4 /*yield*/, web.lists.select("Id").filter("EntityTypeName eq 'AppCatalog'")()];
                    case 2:
                        listId = (_a.sent())[0].Id;
                        return [4 /*yield*/, web.lists.getById(listId).items.filter("AppProductID eq '" + id + "'").top(1)()];
                    case 3:
                        listItems = _a.sent();
                        if (listItems && listItems.length > 0) {
                            appId = listItems[0].Id;
                        }
                        else {
                            throw Error("Did not find the app with id " + id + " in the appcatalog.");
                        }
                        _a.label = 4;
                    case 4:
                        poster = _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(AppCatalog(webUrl, "_api/web/tenantappcatalog/SyncSolutionToTeams(id=" + appId + ")"), "ac.syncSolutionToTeams");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(poster, {})];
                    case 5: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    /**
     * Uploads an app package. Not supported for batching
     *
     * @param filename Filename to create.
     * @param content app package data (eg: the .app or .sppkg file).
     * @param shouldOverWrite Should an app with the same name in the same location be overwritten? (default: true)
     * @returns Promise<IAppAddResult>
     */
    _AppCatalog.prototype.add = function (filename, content, shouldOverWrite) {
        if (shouldOverWrite === void 0) { shouldOverWrite = true; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var adder, r;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        adder = _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(AppCatalog(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__["extractWebUrl"])(this.toUrl()), "_api/web/tenantappcatalog/add(overwrite=" + shouldOverWrite + ",url='" + filename + "')"), "ac.add");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(adder, {
                                body: content, headers: {
                                    "binaryStringRequestBody": "true",
                                },
                            })];
                    case 1:
                        r = _a.sent();
                        return [2 /*return*/, {
                                data: r,
                                file: Object(_files_types__WEBPACK_IMPORTED_MODULE_5__["File"])(Object(_odata__WEBPACK_IMPORTED_MODULE_3__["odataUrlFrom"])(r)),
                            }];
                }
            });
        });
    };
    return _AppCatalog;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var AppCatalog = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_AppCatalog);
var _App = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_App, _super);
    function _App() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * This method deploys an app on the app catalog. It must be called in the context
     * of the tenant app catalog web or it will fail.
     *
     * @param skipFeatureDeployment Deploy the app to the entire tenant
     */
    _App.prototype.deploy = function (skipFeatureDeployment) {
        if (skipFeatureDeployment === void 0) { skipFeatureDeployment = false; }
        return this.do("Deploy(" + skipFeatureDeployment + ")");
    };
    /**
     * This method retracts a deployed app on the app catalog. It must be called in the context
     * of the tenant app catalog web or it will fail.
     */
    _App.prototype.retract = function () {
        return this.do("Retract");
    };
    /**
     * This method allows an app which is already deployed to be installed on a web
     */
    _App.prototype.install = function () {
        return this.do("Install");
    };
    /**
     * This method allows an app which is already installed to be uninstalled on a web
     * Note: when you use the REST API to uninstall a solution package from the site, it is not relocated to the recycle bin
     */
    _App.prototype.uninstall = function () {
        return this.do("Uninstall");
    };
    /**
     * This method allows an app which is already installed to be upgraded on a web
     */
    _App.prototype.upgrade = function () {
        return this.do("Upgrade");
    };
    /**
     * This method removes an app from the app catalog. It must be called in the context
     * of the tenant app catalog web or it will fail.
     */
    _App.prototype.remove = function () {
        return this.do("Remove");
    };
    _App.prototype.do = function (path) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this.clone(App, path));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("app.deploy")
    ], _App.prototype, "deploy", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("app.retract")
    ], _App.prototype, "retract", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("app.install")
    ], _App.prototype, "install", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("app.uninstall")
    ], _App.prototype, "uninstall", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("app.upgrade")
    ], _App.prototype, "upgrade", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("app.remove")
    ], _App.prototype, "remove", null);
    return _App;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var App = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_App);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "At/s":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/funcs.js ***!
  \***********************************************/
/*! exports provided: shareObject, getShareLink, checkPermissions, getSharingInformation, getObjectSharingSettings, unshareObject, deleteLinkByKind, unshareLink, shareWith */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "shareObject", function() { return shareObject; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getShareLink", function() { return getShareLink; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "checkPermissions", function() { return checkPermissions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getSharingInformation", function() { return getSharingInformation; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getObjectSharingSettings", function() { return getObjectSharingSettings; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "unshareObject", function() { return unshareObject; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteLinkByKind", function() { return deleteLinkByKind; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "unshareLink", function() { return unshareLink; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "shareWith", function() { return shareWith; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./types */ "AvPO");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _security_types__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../security/types */ "1dmT");
/* harmony import */ var _splibconfig__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../splibconfig */ "0p/9");











/**
 * Shares an object based on the supplied options
 *
 * @param options The set of options to send to the ShareObject method
 * @param bypass If true any processing is skipped and the options are sent directly to the ShareObject method
 */
function shareObject(o, options, bypass) {
    if (bypass === void 0) { bypass = false; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var roleValue, userStr, postBody;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (bypass) {
                        // if the bypass flag is set send the supplied parameters directly to the service
                        return [2 /*return*/, sendShareObjectRequest(o, options)];
                    }
                    // extend our options with some defaults
                    options = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(options, {
                        group: null,
                        includeAnonymousLinkInEmail: false,
                        propagateAcl: false,
                        useSimplifiedRoles: true,
                    }, true);
                    return [4 /*yield*/, getRoleValue(options.role, options.group)];
                case 1:
                    roleValue = _a.sent();
                    // handle the multiple input types
                    if (!Array.isArray(options.loginNames)) {
                        options.loginNames = [options.loginNames];
                    }
                    userStr = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["jsS"])(options.loginNames.map(function (Key) { return ({ Key: Key }); }));
                    postBody = {
                        peoplePickerInput: userStr,
                        roleValue: roleValue,
                        url: options.url,
                    };
                    if (options.emailData !== undefined && options.emailData !== null) {
                        postBody = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(postBody, {
                            emailBody: options.emailData.body,
                            emailSubject: options.emailData.subject !== undefined ? options.emailData.subject : "Shared with you.",
                            sendEmail: true,
                        });
                    }
                    return [2 /*return*/, sendShareObjectRequest(o, postBody)];
            }
        });
    });
}
/**
 * Gets a sharing link for the supplied
 *
 * @param kind The kind of link to share
 * @param expiration The optional expiration for this link
 */
function getShareLink(kind, expiration) {
    if (expiration === void 0) { expiration = null; }
    // date needs to be an ISO string or null
    var expString = expiration !== null ? expiration.toISOString() : null;
    // clone using the factory and send the request
    var o = _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "shareLink"), "sh.getShareLink");
    return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(o, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({
        request: {
            createLink: true,
            emailData: null,
            settings: {
                expiration: expString,
                linkKind: kind,
            },
        },
    }));
}
/**
 * Checks Permissions on the list of Users and returns back role the users have on the Item.
 *
 * @param recipients The array of Entities for which Permissions need to be checked.
 */
function checkPermissions(recipients) {
    var o = _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "checkPermissions"), "sh.checkPermissions");
    return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(o, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ recipients: recipients }));
}
/**
 * Get Sharing Information.
 *
 * @param request The SharingInformationRequest Object.
 * @param expands Expand more fields.
 *
 */
function getSharingInformation(request, expands) {
    if (request === void 0) { request = null; }
    if (expands === void 0) { expands = []; }
    var o = _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "getSharingInformation"), "sh.getSharingInformation");
    return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(o.expand.apply(o, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])(expands)), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ request: request }));
}
/**
 * Gets the sharing settings of an item.
 *
 * @param useSimplifiedRoles Determines whether to use simplified roles.
 */
function getObjectSharingSettings(useSimplifiedRoles) {
    if (useSimplifiedRoles === void 0) { useSimplifiedRoles = true; }
    var o = _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "getObjectSharingSettings"), "sh.getObjectSharingSettings");
    return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(o, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ useSimplifiedRoles: useSimplifiedRoles }));
}
/**
 * Unshares this object
 */
function unshareObject() {
    return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "unshareObject"), "sh.unshareObject"));
}
/**
 * Deletes a link by type
 *
 * @param kind Deletes a sharing link by the kind of link
 */
function deleteLinkByKind(linkKind) {
    return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "deleteLinkByKind"), "sh.deleteLinkByKind"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ linkKind: linkKind }));
}
/**
 * Removes the specified link to the item.
 *
 * @param kind The kind of link to be deleted.
 * @param shareId
 */
function unshareLink(linkKind, shareId) {
    if (shareId === void 0) { shareId = _splibconfig__WEBPACK_IMPORTED_MODULE_10__["emptyGuid"]; }
    return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "unshareLink"), "sh.unshareLink"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ linkKind: linkKind, shareId: shareId }));
}
/**
 * Shares this instance with the supplied users
 *
 * @param loginNames Resolved login names to share
 * @param role The role
 * @param requireSignin True to require the user is authenticated, otherwise false
 * @param propagateAcl True to apply this share to all children
 * @param emailData If supplied an email will be sent with the indicated properties
 */
function shareWith(o, loginNames, role, requireSignin, propagateAcl, emailData) {
    if (requireSignin === void 0) { requireSignin = false; }
    if (propagateAcl === void 0) { propagateAcl = false; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var userStr, roleFilter, w, def, postBody;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    // handle the multiple input types
                    if (!Array.isArray(loginNames)) {
                        loginNames = [loginNames];
                    }
                    userStr = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["jsS"])(loginNames.map(function (login) { return { Key: login }; }));
                    roleFilter = role === _types__WEBPACK_IMPORTED_MODULE_6__["SharingRole"].Edit ? _types__WEBPACK_IMPORTED_MODULE_6__["RoleType"].Contributor : _types__WEBPACK_IMPORTED_MODULE_6__["RoleType"].Reader;
                    w = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableCollection"])("_api/web", "roledefinitions");
                    return [4 /*yield*/, w.select("Id").filter("RoleTypeKind eq " + roleFilter).get()];
                case 1:
                    def = _a.sent();
                    if (!Array.isArray(def) || def.length < 1) {
                        throw Error("Could not locate a role defintion with RoleTypeKind " + roleFilter);
                    }
                    postBody = {
                        includeAnonymousLinkInEmail: requireSignin,
                        peoplePickerInput: userStr,
                        propagateAcl: propagateAcl,
                        roleValue: "role:" + def[0].Id,
                        useSimplifiedRoles: true,
                    };
                    if (emailData !== undefined) {
                        postBody = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(postBody, {
                            emailBody: emailData.body,
                            emailSubject: emailData.subject !== undefined ? emailData.subject : "",
                            sendEmail: true,
                        });
                    }
                    return [2 /*return*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(o.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "shareObject"), "sh.shareWith"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(postBody))];
            }
        });
    });
}
function sendShareObjectRequest(o, options) {
    var w = _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(Object(_webs_types__WEBPACK_IMPORTED_MODULE_5__["Web"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__["extractWebUrl"])(o.toUrl()), "/_api/SP.Web.ShareObject"), "sh.sendShareObjectRequest");
    return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(w.expand("UsersWithAccessRequests", "GroupsSharedWith"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(options));
}
/**
 * Calculates the roleValue string used in the sharing query
 *
 * @param role The Sharing Role
 * @param group The Group type
 */
function getRoleValue(role, group) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var _a, g1, g2, roleFilter, def;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_b) {
            switch (_b.label) {
                case 0:
                    if (!(group !== undefined && group !== null)) return [3 /*break*/, 7];
                    _a = group;
                    switch (_a) {
                        case _types__WEBPACK_IMPORTED_MODULE_6__["RoleType"].Contributor: return [3 /*break*/, 1];
                        case _types__WEBPACK_IMPORTED_MODULE_6__["RoleType"].Reader: return [3 /*break*/, 3];
                        case _types__WEBPACK_IMPORTED_MODULE_6__["RoleType"].Guest: return [3 /*break*/, 3];
                    }
                    return [3 /*break*/, 5];
                case 1: return [4 /*yield*/, Object(_webs_types__WEBPACK_IMPORTED_MODULE_5__["Web"])("_api/web", "associatedmembergroup").select("Id")()];
                case 2:
                    g1 = _b.sent();
                    return [2 /*return*/, "group: " + g1.Id];
                case 3: return [4 /*yield*/, Object(_webs_types__WEBPACK_IMPORTED_MODULE_5__["Web"])("_api/web", "associatedvisitorgroup").select("Id")()];
                case 4:
                    g2 = _b.sent();
                    return [2 /*return*/, "group: " + g2.Id];
                case 5: throw Error("Could not determine role value for supplied value. Contributor, Reader, and Guest are supported");
                case 6: return [3 /*break*/, 9];
                case 7:
                    roleFilter = role === _types__WEBPACK_IMPORTED_MODULE_6__["SharingRole"].Edit ? _types__WEBPACK_IMPORTED_MODULE_6__["RoleType"].Contributor : _types__WEBPACK_IMPORTED_MODULE_6__["RoleType"].Reader;
                    return [4 /*yield*/, Object(_security_types__WEBPACK_IMPORTED_MODULE_9__["RoleDefinitions"])("_api/web").select("Id").top(1).filter("RoleTypeKind eq " + roleFilter)()];
                case 8:
                    def = _b.sent();
                    if (def.length < 1) {
                        throw Error("Could not locate associated role definition for supplied role. Edit and View are supported");
                    }
                    return [2 /*return*/, "role: " + def[0].Id];
                case 9: return [2 /*return*/];
            }
        });
    });
}
//# sourceMappingURL=funcs.js.map

/***/ }),

/***/ "AvPO":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/types.js ***!
  \***********************************************/
/*! exports provided: SharingRole, SPSharedObjectType, SharingDomainRestrictionMode, SharingOperationStatusCode, SharingLinkKind, RoleType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SharingRole", function() { return SharingRole; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SPSharedObjectType", function() { return SPSharedObjectType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SharingDomainRestrictionMode", function() { return SharingDomainRestrictionMode; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SharingOperationStatusCode", function() { return SharingOperationStatusCode; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SharingLinkKind", function() { return SharingLinkKind; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RoleType", function() { return RoleType; });
/**
 * Indicates the role of the sharing link
 */
var SharingRole;
(function (SharingRole) {
    SharingRole[SharingRole["None"] = 0] = "None";
    SharingRole[SharingRole["View"] = 1] = "View";
    SharingRole[SharingRole["Edit"] = 2] = "Edit";
    SharingRole[SharingRole["Owner"] = 3] = "Owner";
})(SharingRole || (SharingRole = {}));
var SPSharedObjectType;
(function (SPSharedObjectType) {
    SPSharedObjectType[SPSharedObjectType["Unknown"] = 0] = "Unknown";
    SPSharedObjectType[SPSharedObjectType["File"] = 1] = "File";
    SPSharedObjectType[SPSharedObjectType["Folder"] = 2] = "Folder";
    SPSharedObjectType[SPSharedObjectType["Item"] = 3] = "Item";
    SPSharedObjectType[SPSharedObjectType["List"] = 4] = "List";
    SPSharedObjectType[SPSharedObjectType["Web"] = 5] = "Web";
    SPSharedObjectType[SPSharedObjectType["Max"] = 6] = "Max";
})(SPSharedObjectType || (SPSharedObjectType = {}));
var SharingDomainRestrictionMode;
(function (SharingDomainRestrictionMode) {
    SharingDomainRestrictionMode[SharingDomainRestrictionMode["None"] = 0] = "None";
    SharingDomainRestrictionMode[SharingDomainRestrictionMode["AllowList"] = 1] = "AllowList";
    SharingDomainRestrictionMode[SharingDomainRestrictionMode["BlockList"] = 2] = "BlockList";
})(SharingDomainRestrictionMode || (SharingDomainRestrictionMode = {}));
var SharingOperationStatusCode;
(function (SharingOperationStatusCode) {
    /**
     * The share operation completed without errors.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["CompletedSuccessfully"] = 0] = "CompletedSuccessfully";
    /**
     * The share operation completed and generated requests for access.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["AccessRequestsQueued"] = 1] = "AccessRequestsQueued";
    /**
     * The share operation failed as there were no resolved users.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["NoResolvedUsers"] = -1] = "NoResolvedUsers";
    /**
     * The share operation failed due to insufficient permissions.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["AccessDenied"] = -2] = "AccessDenied";
    /**
     * The share operation failed when attempting a cross site share, which is not supported.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["CrossSiteRequestNotSupported"] = -3] = "CrossSiteRequestNotSupported";
    /**
     * The sharing operation failed due to an unknown error.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["UnknowError"] = -4] = "UnknowError";
    /**
     * The text you typed is too long. Please shorten it.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["EmailBodyTooLong"] = -5] = "EmailBodyTooLong";
    /**
     * The maximum number of unique scopes in the list has been exceeded.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["ListUniqueScopesExceeded"] = -6] = "ListUniqueScopesExceeded";
    /**
     * The share operation failed because a sharing capability is disabled in the site.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["CapabilityDisabled"] = -7] = "CapabilityDisabled";
    /**
     * The specified object for the share operation is not supported.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["ObjectNotSupported"] = -8] = "ObjectNotSupported";
    /**
     * A SharePoint group cannot contain another SharePoint group.
     */
    SharingOperationStatusCode[SharingOperationStatusCode["NestedGroupsNotSupported"] = -9] = "NestedGroupsNotSupported";
})(SharingOperationStatusCode || (SharingOperationStatusCode = {}));
var SharingLinkKind;
(function (SharingLinkKind) {
    /**
     * Uninitialized link
     */
    SharingLinkKind[SharingLinkKind["Uninitialized"] = 0] = "Uninitialized";
    /**
     * Direct link to the object being shared
     */
    SharingLinkKind[SharingLinkKind["Direct"] = 1] = "Direct";
    /**
     * Organization-shareable link to the object being shared with view permissions
     */
    SharingLinkKind[SharingLinkKind["OrganizationView"] = 2] = "OrganizationView";
    /**
     * Organization-shareable link to the object being shared with edit permissions
     */
    SharingLinkKind[SharingLinkKind["OrganizationEdit"] = 3] = "OrganizationEdit";
    /**
     * View only anonymous link
     */
    SharingLinkKind[SharingLinkKind["AnonymousView"] = 4] = "AnonymousView";
    /**
     * Read/Write anonymous link
     */
    SharingLinkKind[SharingLinkKind["AnonymousEdit"] = 5] = "AnonymousEdit";
    /**
     * Flexible sharing Link where properties can change without affecting link URL
     */
    SharingLinkKind[SharingLinkKind["Flexible"] = 6] = "Flexible";
})(SharingLinkKind || (SharingLinkKind = {}));
var RoleType;
(function (RoleType) {
    RoleType[RoleType["None"] = 0] = "None";
    RoleType[RoleType["Guest"] = 1] = "Guest";
    RoleType[RoleType["Reader"] = 2] = "Reader";
    RoleType[RoleType["Contributor"] = 3] = "Contributor";
    RoleType[RoleType["WebDesigner"] = 4] = "WebDesigner";
    RoleType[RoleType["Administrator"] = 5] = "Administrator";
})(RoleType || (RoleType = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "B3L+":
/*!********************************************!*\
  !*** ./node_modules/@pnp/logging/index.js ***!
  \********************************************/
/*! exports provided: Logger, LogLevel, ConsoleListener, FunctionListener */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _logger__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./logger */ "Vpmm");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Logger", function() { return _logger__WEBPACK_IMPORTED_MODULE_0__["Logger"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "LogLevel", function() { return _logger__WEBPACK_IMPORTED_MODULE_0__["LogLevel"]; });

/* harmony import */ var _listeners__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./listeners */ "aH1g");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ConsoleListener", function() { return _listeners__WEBPACK_IMPORTED_MODULE_1__["ConsoleListener"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FunctionListener", function() { return _listeners__WEBPACK_IMPORTED_MODULE_1__["FunctionListener"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "BK4w":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/navigation/types.js ***!
  \**************************************************/
/*! exports provided: _NavigationNodes, NavigationNodes, _NavigationNode, NavigationNode, _Navigation, Navigation, _NavigationService, NavigationService */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_NavigationNodes", function() { return _NavigationNodes; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "NavigationNodes", function() { return NavigationNodes; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_NavigationNode", function() { return _NavigationNode; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "NavigationNode", function() { return NavigationNode; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Navigation", function() { return _Navigation; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Navigation", function() { return Navigation; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_NavigationService", function() { return _NavigationService; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "NavigationService", function() { return NavigationService; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../telemetry */ "h3gW");








/**
 * Represents a collection of navigation nodes
 *
 */
var _NavigationNodes = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_NavigationNodes, _super);
    function _NavigationNodes() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a navigation node by id
     *
     * @param id The id of the node
     */
    _NavigationNodes.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].configure(NavigationNode(this).concat("(" + id + ")"), "nns.getById");
    };
    /**
     * Adds a new node to the collection
     *
     * @param title Display name of the node
     * @param url The url of the node
     * @param visible If true the node is visible, otherwise it is hidden (default: true)
     */
    _NavigationNodes.prototype.add = function (title, url, visible) {
        if (visible === void 0) { visible = true; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])("SP.NavigationNode"), {
                            IsVisible: visible,
                            Title: title,
                            Url: url,
                        }));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(NavigationNodes, null), postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                node: this.getById(data.Id),
                            }];
                }
            });
        });
    };
    /**
     * Moves a node to be after another node in the navigation
     *
     * @param nodeId Id of the node to move
     * @param previousNodeId Id of the node after which we move the node specified by nodeId
     */
    _NavigationNodes.prototype.moveAfter = function (nodeId, previousNodeId) {
        var postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({
            nodeId: nodeId,
            previousNodeId: previousNodeId,
        });
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(NavigationNodes, "MoveAfter"), postBody);
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("nns.add")
    ], _NavigationNodes.prototype, "add", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("nns.moveAfter")
    ], _NavigationNodes.prototype, "moveAfter", null);
    return _NavigationNodes;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var NavigationNodes = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_NavigationNodes);
/**
 * Represents an instance of a navigation node
 *
 */
var _NavigationNode = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_NavigationNode, _super);
    function _NavigationNode() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["deleteable"])("nn");
        return _this;
    }
    Object.defineProperty(_NavigationNode.prototype, "children", {
        /**
         * Represents the child nodes of this node
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].configure(NavigationNodes(this, "children"), "nn.children");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Updates this node
     *
     * @param properties Properties used to update this node
     */
    _NavigationNode.prototype.update = function (properties) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])("SP.NavigationNode"), properties), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["headers"])({ "X-HTTP-Method": "MERGE" }));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this, postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                node: this,
                            }];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("nn.update")
    ], _NavigationNode.prototype, "update", null);
    return _NavigationNode;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var NavigationNode = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_NavigationNode);
/**
 * Exposes the navigation components
 *
 */
var _Navigation = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Navigation, _super);
    function _Navigation() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(_Navigation.prototype, "quicklaunch", {
        /**
         * Gets the quicklaunch navigation nodes for the current context
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].configure(NavigationNodes(this, "quicklaunch"), "n.quicklaunch");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Navigation.prototype, "topNavigationBar", {
        /**
         * Gets the top bar navigation nodes for the current context
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].configure(NavigationNodes(this, "topnavigationbar"), "n.topnavigationbar");
        },
        enumerable: true,
        configurable: true
    });
    _Navigation = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_5__["defaultPath"])("navigation")
    ], _Navigation);
    return _Navigation;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryable"]));

var Navigation = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Navigation);
/**
 * Represents the top level navigation service
 */
var _NavigationService = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_NavigationService, _super);
    function _NavigationService(path) {
        if (path === void 0) { path = null; }
        return _super.call(this, "_api/navigation", path) || this;
    }
    /**
     * The MenuState service operation returns a Menu-State (dump) of a SiteMapProvider on a site.
     *
     * @param menuNodeKey MenuNode.Key of the start node within the SiteMapProvider If no key is provided the SiteMapProvider.RootNode will be the root of the menu state.
     * @param depth Depth of the dump. If no value is provided a dump with the depth of 10 is returned
     * @param mapProviderName The name identifying the SiteMapProvider to be used
     * @param customProperties comma seperated list of custom properties to be returned.
     */
    _NavigationService.prototype.getMenuState = function (menuNodeKey, depth, mapProviderName, customProperties) {
        if (menuNodeKey === void 0) { menuNodeKey = null; }
        if (depth === void 0) { depth = 10; }
        if (mapProviderName === void 0) { mapProviderName = null; }
        if (customProperties === void 0) { customProperties = null; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(NavigationService("MenuState"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({
            customProperties: customProperties,
            depth: depth,
            mapProviderName: mapProviderName,
            menuNodeKey: menuNodeKey,
        }));
    };
    /**
     * Tries to get a SiteMapNode.Key for a given URL within a site collection.
     *
     * @param currentUrl A url representing the SiteMapNode
     * @param mapProviderName The name identifying the SiteMapProvider to be used
     */
    _NavigationService.prototype.getMenuNodeKey = function (currentUrl, mapProviderName) {
        if (mapProviderName === void 0) { mapProviderName = null; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(NavigationService("MenuNodeKey"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({
            currentUrl: currentUrl,
            mapProviderName: mapProviderName,
        }));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("ns.getMenuState")
    ], _NavigationService.prototype, "getMenuState", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("ns.getMenuNodeKey")
    ], _NavigationService.prototype, "getMenuNodeKey", null);
    return _NavigationService;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryable"]));

var NavigationService = function (path) { return new _NavigationService(path); };
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "Bwa7":
/*!*******************************************!*\
  !*** ./node_modules/@pnp/sp/lists/web.js ***!
  \*******************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "hy0S");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");







Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "lists", _types__WEBPACK_IMPORTED_MODULE_3__["Lists"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "siteUserInfoList", _types__WEBPACK_IMPORTED_MODULE_3__["List"], "siteuserinfolist");
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "defaultDocumentLibrary", _types__WEBPACK_IMPORTED_MODULE_3__["List"], "DefaultDocumentLibrary");
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "customListTemplates", _sharepointqueryable__WEBPACK_IMPORTED_MODULE_5__["SharePointQueryableCollection"], "getcustomlisttemplates");
_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"].prototype.getList = function (listRelativeUrl) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_3__["List"])(this, "getList('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_6__["escapeQueryStrValue"])(listRelativeUrl) + "')");
};
_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"].prototype.getCatalog = function (type) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var data;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, this.clone(_webs_types__WEBPACK_IMPORTED_MODULE_2__["Web"], "getcatalog(" + type + ")").select("Id").get()];
                case 1:
                    data = _a.sent();
                    return [2 /*return*/, Object(_types__WEBPACK_IMPORTED_MODULE_3__["List"])(Object(_odata__WEBPACK_IMPORTED_MODULE_4__["odataUrlFrom"])(data))];
            }
        });
    });
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "CKSE":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/features/web.js ***!
  \**********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "HHhI");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "features", _types__WEBPACK_IMPORTED_MODULE_2__["Features"]);
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "CbhW":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/features/index.js ***!
  \************************************************/
/*! exports provided: Feature, Features */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _site__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./site */ "Zlzk");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./web */ "CKSE");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "HHhI");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Feature", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["Feature"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Features", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["Features"]; });




//# sourceMappingURL=index.js.map

/***/ }),

/***/ "D2kI":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/search/index.js ***!
  \**********************************************/
/*! exports provided: SearchQueryBuilder, SearchResults, Search, Suggest, SortDirection, ReorderingRuleMatchType, QueryPropertyValueType, SearchBuiltInSourceId */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _query__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./query */ "MkPD");
/* harmony import */ var _suggest__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./suggest */ "fB9m");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "YpSs");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SortDirection", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["SortDirection"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ReorderingRuleMatchType", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["ReorderingRuleMatchType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "QueryPropertyValueType", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["QueryPropertyValueType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SearchBuiltInSourceId", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["SearchBuiltInSourceId"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SearchQueryBuilder", function() { return _query__WEBPACK_IMPORTED_MODULE_1__["SearchQueryBuilder"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SearchResults", function() { return _query__WEBPACK_IMPORTED_MODULE_1__["SearchResults"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Search", function() { return _query__WEBPACK_IMPORTED_MODULE_1__["Search"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Suggest", function() { return _suggest__WEBPACK_IMPORTED_MODULE_2__["Suggest"]; });








_rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"].prototype.search = function (query) {
    return Object(_query__WEBPACK_IMPORTED_MODULE_1__["Search"])(this._baseUrl, this._options)(query);
};
_rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"].prototype.searchWithCaching = function (query, options) {
    return (new _query__WEBPACK_IMPORTED_MODULE_1__["_Search"](this._baseUrl)).configure(this._options).usingCaching(options).execute(query);
};
_rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"].prototype.searchSuggest = function (query) {
    return Object(_suggest__WEBPACK_IMPORTED_MODULE_2__["Suggest"])(this._baseUrl, this._options)(typeof query === "string" ? { querytext: query } : query);
};
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "DL6e":
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/field.js ***!
  \*********************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _fields_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../fields/types */ "KYlJ");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./funcs */ "T2j4");


_fields_types__WEBPACK_IMPORTED_MODULE_0__["_Field"].prototype.titleResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_1__["getValueForUICultureBinder"])("titleResource");
_fields_types__WEBPACK_IMPORTED_MODULE_0__["_Field"].prototype.descriptionResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_1__["getValueForUICultureBinder"])("descriptionResource");
//# sourceMappingURL=field.js.map

/***/ }),

/***/ "DqvY":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/sphttpclient.js ***!
  \**********************************************/
/*! exports provided: SPHttpClient */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SPHttpClient", function() { return SPHttpClient; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _splibconfig__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./splibconfig */ "0p/9");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./utils/extractweburl */ "tSEd");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./telemetry */ "h3gW");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/odata */ "xT8B");






var SPHttpClient = /** @class */ (function () {
    function SPHttpClient(_impl) {
        if (_impl === void 0) { _impl = _splibconfig__WEBPACK_IMPORTED_MODULE_2__["SPRuntimeConfig"].fetchClientFactory(); }
        this._impl = _impl;
        this._digestCache = getDigestFactory(this);
    }
    SPHttpClient.prototype.fetch = function (url, options) {
        if (options === void 0) { options = {}; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var opts, headers, methodName, clientTag, digest;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        opts = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(options, { cache: "no-cache", credentials: "same-origin" }, true);
                        headers = new Headers();
                        // first we add the global headers so they can be overwritten by any passed in locally to this call
                        Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["mergeHeaders"])(headers, _splibconfig__WEBPACK_IMPORTED_MODULE_2__["SPRuntimeConfig"].headers);
                        // second we add the local options so we can overwrite the globals
                        Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["mergeHeaders"])(headers, options.headers);
                        // lastly we apply any default headers we need that may not exist
                        if (!headers.has("Accept")) {
                            headers.append("Accept", "application/json");
                        }
                        if (!headers.has("Content-Type")) {
                            headers.append("Content-Type", "application/json;odata=verbose;charset=utf-8");
                        }
                        if (!headers.has("X-ClientService-ClientTag")) {
                            methodName = _telemetry__WEBPACK_IMPORTED_MODULE_4__["tag"].getClientTag(headers);
                            clientTag = "PnPCoreJS:2.0.5:" + methodName;
                            if (clientTag.length > 32) {
                                clientTag = clientTag.substr(0, 32);
                            }
                            headers.append("X-ClientService-ClientTag", clientTag);
                        }
                        opts = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(opts, { headers: headers });
                        if (!(opts.method && opts.method.toUpperCase() !== "GET" && !headers.has("X-RequestDigest") && !headers.has("Authorization"))) return [3 /*break*/, 2];
                        return [4 /*yield*/, this._digestCache(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_3__["extractWebUrl"])(url))];
                    case 1:
                        digest = _a.sent();
                        headers.append("X-RequestDigest", digest);
                        _a.label = 2;
                    case 2: return [2 /*return*/, this.fetchRaw(url, opts)];
                }
            });
        });
    };
    SPHttpClient.prototype.fetchRaw = function (url, options) {
        var _this = this;
        if (options === void 0) { options = {}; }
        // here we need to normalize the headers
        var rawHeaders = new Headers();
        Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["mergeHeaders"])(rawHeaders, options.headers);
        options = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(options, { headers: rawHeaders });
        var retry = function (ctx) {
            // handles setting the proper timeout for a retry
            var setRetry = function (response) {
                var delay;
                if (response.headers.has("Retry-After")) {
                    // if we have gotten a header, use that value as the delay value in seconds
                    delay = parseInt(response.headers.get("Retry-After"), 10) * 1000;
                }
                else {
                    // grab our current delay
                    delay = ctx.delay;
                    // Increment our counters.
                    ctx.delay *= 2;
                }
                ctx.attempts++;
                // If we have exceeded the retry count, reject.
                if (ctx.retryCount <= ctx.attempts) {
                    ctx.reject(Error("Retry count exceeded (" + ctx.retryCount + ") for request. Response status: [" + response.status + "] " + response.statusText));
                }
                else {
                    // Set our retry timeout for {delay} milliseconds.
                    setTimeout(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["getCtxCallback"])(_this, retry, ctx), delay);
                }
            };
            // send the actual request
            _this._impl.fetch(url, options).then(function (response) {
                if (response.status === 429) {
                    // we have been throttled
                    setRetry(response);
                }
                else {
                    ctx.resolve(response);
                }
            }).catch(function (response) {
                if (response.status === 503 || response.status === 504) {
                    // http status code 503 or 504, we can retry this
                    setRetry(response);
                }
                else {
                    ctx.reject(response);
                }
            });
        };
        return new Promise(function (resolve, reject) {
            retry.call(_this, {
                attempts: 0,
                delay: 100,
                reject: reject,
                resolve: resolve,
                retryCount: 7,
            });
        });
    };
    SPHttpClient.prototype.get = function (url, options) {
        if (options === void 0) { options = {}; }
        var opts = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(options, { method: "GET" });
        return this.fetch(url, opts);
    };
    SPHttpClient.prototype.post = function (url, options) {
        if (options === void 0) { options = {}; }
        var opts = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(options, { method: "POST" });
        return this.fetch(url, opts);
    };
    SPHttpClient.prototype.patch = function (url, options) {
        if (options === void 0) { options = {}; }
        var opts = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(options, { method: "PATCH" });
        return this.fetch(url, opts);
    };
    SPHttpClient.prototype.delete = function (url, options) {
        if (options === void 0) { options = {}; }
        var opts = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(options, { method: "DELETE" });
        return this.fetch(url, opts);
    };
    return SPHttpClient;
}());

// allows for the caching of digests across all HttpClient's which each have their own DigestCache wrapper.
var digests = new Map();
function getDigestFactory(client) {
    var _this = this;
    return function (webUrl) { return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(_this, void 0, void 0, function () {
        var cachedDigest, now, url, headers, resp, parsed, newCachedDigest;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    cachedDigest = digests.get(webUrl);
                    if (cachedDigest !== undefined) {
                        now = new Date();
                        if (now < cachedDigest.expiration) {
                            return [2 /*return*/, cachedDigest.value];
                        }
                    }
                    url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"])(webUrl, "/_api/contextinfo");
                    headers = {
                        "Accept": "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose;charset=utf-8",
                    };
                    return [4 /*yield*/, client.fetchRaw(url, {
                            cache: "no-cache",
                            credentials: "same-origin",
                            headers: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(headers, _splibconfig__WEBPACK_IMPORTED_MODULE_2__["SPRuntimeConfig"].headers, true),
                            method: "POST",
                        })];
                case 1:
                    resp = _a.sent();
                    return [4 /*yield*/, (new _pnp_odata__WEBPACK_IMPORTED_MODULE_5__["ODataParser"]()).parse(resp).then(function (r) { return r.GetContextWebInformation; })];
                case 2:
                    parsed = _a.sent();
                    newCachedDigest = {
                        expiration: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["dateAdd"])(new Date(), "second", parsed.FormDigestTimeoutSeconds),
                        value: parsed.FormDigestValue,
                    };
                    digests.set(webUrl, newCachedDigest);
                    return [2 /*return*/, newCachedDigest.value];
            }
        });
    }); };
}
//# sourceMappingURL=sphttpclient.js.map

/***/ }),

/***/ "E/FW":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/social/index.js ***!
  \**********************************************/
/*! exports provided: MySocial, Social, SocialActorType, SocialActorTypes, SocialFollowResult, SocialStatusCode */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./types */ "6a65");
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "MySocial", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["MySocial"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Social", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["Social"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SocialActorType", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["SocialActorType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SocialActorTypes", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["SocialActorTypes"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SocialFollowResult", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["SocialFollowResult"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SocialStatusCode", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["SocialStatusCode"]; });




Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_1__["SPRest"].prototype, "social", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_0__["Social"])(this._baseUrl);
    },
});
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "E6kk":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/sharepointqueryable.js ***!
  \*****************************************************/
/*! exports provided: spInvokableFactory, _SharePointQueryable, SharePointQueryable, _SharePointQueryableCollection, SharePointQueryableCollection, _SharePointQueryableInstance, SharePointQueryableInstance, deleteable, deleteableWithETag */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spInvokableFactory", function() { return spInvokableFactory; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SharePointQueryable", function() { return _SharePointQueryable; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryable", function() { return SharePointQueryable; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SharePointQueryableCollection", function() { return _SharePointQueryableCollection; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryableCollection", function() { return SharePointQueryableCollection; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SharePointQueryableInstance", function() { return _SharePointQueryableInstance; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryableInstance", function() { return SharePointQueryableInstance; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteable", function() { return deleteable; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "deleteableWithETag", function() { return deleteableWithETag; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _pnp_logging__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/logging */ "B3L+");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./utils/metadata */ "z067");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./telemetry */ "h3gW");







var spInvokableFactory = function (f) {
    return Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["invokableFactory"])(f);
};
/**
 * SharePointQueryable Base Class
 *
 */
var _SharePointQueryable = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SharePointQueryable, _super);
    /**
     * Creates a new instance of the SharePointQueryable class
     *
     * @constructor
     * @param baseUrl A string or SharePointQueryable that should form the base part of the url
     *
     */
    function _SharePointQueryable(baseUrl, path) {
        var _this = this;
        var url = "";
        var parentUrl = "";
        var query = new Map();
        if (typeof baseUrl === "string") {
            // we need to do some extra parsing to get the parent url correct if we are
            // being created from just a string.
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isUrlAbsolute"])(baseUrl) || baseUrl.lastIndexOf("/") < 0) {
                parentUrl = baseUrl;
                url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"])(baseUrl, path);
            }
            else if (baseUrl.lastIndexOf("/") > baseUrl.lastIndexOf("(")) {
                // .../items(19)/fields
                var index = baseUrl.lastIndexOf("/");
                parentUrl = baseUrl.slice(0, index);
                path = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"])(baseUrl.slice(index), path);
                url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"])(parentUrl, path);
            }
            else {
                // .../items(19)
                var index = baseUrl.lastIndexOf("(");
                parentUrl = baseUrl.slice(0, index);
                url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"])(baseUrl, path);
            }
        }
        else {
            parentUrl = baseUrl.toUrl();
            url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"])(parentUrl, path || "");
            var target = baseUrl.query.get("@target");
            if (target !== undefined) {
                query.set("@target", target);
            }
        }
        // init base with correct values for data seed
        _this = _super.call(this, {
            parentUrl: parentUrl,
            query: query,
            url: url,
        }) || this;
        // post init actions
        if (typeof baseUrl !== "string") {
            _this.configureFrom(baseUrl);
        }
        _this._forceCaching = false;
        return _this;
    }
    /**
     * Gets the full url with query information
     */
    _SharePointQueryable.prototype.toUrlAndQuery = function () {
        var aliasedParams = new Map(this.query);
        var url = this.toUrl().replace(/'!(@.*?)::(.*?)'/ig, function (match, labelName, value) {
            _pnp_logging__WEBPACK_IMPORTED_MODULE_3__["Logger"].write("Rewriting aliased parameter from match " + match + " to label: " + labelName + " value: " + value, 0 /* Verbose */);
            aliasedParams.set(labelName, "'" + value + "'");
            return labelName;
        });
        if (aliasedParams.size > 0) {
            var char = url.indexOf("?") > -1 ? "&" : "?";
            url += "" + char + Array.from(aliasedParams).map(function (v) { return v[0] + "=" + v[1]; }).join("&");
        }
        return url;
    };
    /**
     * Choose which fields to return
     *
     * @param selects One or more fields to return
     */
    _SharePointQueryable.prototype.select = function () {
        var selects = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            selects[_i] = arguments[_i];
        }
        if (selects.length > 0) {
            this.query.set("$select", selects.join(","));
        }
        return this;
    };
    _SharePointQueryable.prototype.get = function (options) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spGet"])(this, options);
    };
    /**
     * Expands fields such as lookups to get additional data
     *
     * @param expands The Fields for which to expand the values
     */
    _SharePointQueryable.prototype.expand = function () {
        var expands = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            expands[_i] = arguments[_i];
        }
        if (expands.length > 0) {
            this.query.set("$expand", expands.join(","));
        }
        return this;
    };
    /**
     * Clones this SharePointQueryable into a new SharePointQueryable instance of T
     * @param factory Constructor used to create the new instance
     * @param additionalPath Any additional path to include in the clone
     * @param includeBatch If true this instance's batch will be added to the cloned instance
     */
    _SharePointQueryable.prototype.clone = function (factory, additionalPath, includeBatch) {
        if (includeBatch === void 0) { includeBatch = true; }
        var clone = _super.prototype.cloneTo.call(this, factory(this, additionalPath), { includeBatch: includeBatch });
        // handle sp specific clone actions
        var t = "@target";
        if (this.query.has(t)) {
            clone.query.set(t, this.query.get(t));
        }
        return clone;
    };
    /**
     * The default action for this object (unless overridden spGet)
     *
     * @param options optional request options
     */
    _SharePointQueryable.prototype.defaultAction = function (options) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spGet"])(this, options);
    };
    /**
     * Gets a parent for this instance as specified
     *
     * @param factory The contructor for the class to create
     */
    _SharePointQueryable.prototype.getParent = function (factory, baseUrl, path, batch) {
        if (baseUrl === void 0) { baseUrl = this.parentUrl; }
        var parent = factory(baseUrl, path).configureFrom(this);
        var t = "@target";
        if (this.query.has(t)) {
            parent.query.set(t, this.query.get(t));
        }
        if (batch !== undefined) {
            parent = parent.inBatch(batch);
        }
        return parent;
    };
    return _SharePointQueryable;
}(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["Queryable"]));

var SharePointQueryable = spInvokableFactory(_SharePointQueryable);
/**
 * Represents a REST collection which can be filtered, paged, and selected
 *
 */
var _SharePointQueryableCollection = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SharePointQueryableCollection, _super);
    function _SharePointQueryableCollection() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Filters the returned collection (https://msdn.microsoft.com/en-us/library/office/fp142385.aspx#bk_supported)
     *
     * @param filter The string representing the filter query
     */
    _SharePointQueryableCollection.prototype.filter = function (filter) {
        this.query.set("$filter", filter);
        return this;
    };
    /**
     * Orders based on the supplied fields
     *
     * @param orderby The name of the field on which to sort
     * @param ascending If false DESC is appended, otherwise ASC (default)
     */
    _SharePointQueryableCollection.prototype.orderBy = function (orderBy, ascending) {
        if (ascending === void 0) { ascending = true; }
        var o = "$orderby";
        var query = this.query.has(o) ? this.query.get(o).split(",") : [];
        query.push(orderBy + " " + (ascending ? "asc" : "desc"));
        this.query.set(o, query.join(","));
        return this;
    };
    /**
     * Skips the specified number of items
     *
     * @param skip The number of items to skip
     */
    _SharePointQueryableCollection.prototype.skip = function (skip) {
        this.query.set("$skip", skip.toString());
        return this;
    };
    /**
     * Limits the query to only return the specified number of items
     *
     * @param top The query row limit
     */
    _SharePointQueryableCollection.prototype.top = function (top) {
        this.query.set("$top", top.toString());
        return this;
    };
    return _SharePointQueryableCollection;
}(_SharePointQueryable));

var SharePointQueryableCollection = spInvokableFactory(_SharePointQueryableCollection);
/**
 * Represents an instance that can be selected
 *
 */
var _SharePointQueryableInstance = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SharePointQueryableInstance, _super);
    function _SharePointQueryableInstance() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Curries the update function into the common pieces
     *
     * @param type
     * @param mapper
     */
    _SharePointQueryableInstance.prototype._update = function (type, mapper) {
        var _this = this;
        return function (props) { return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(_this, type + ".Update"), {
            body: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["jsS"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_4__["metadata"])(type), props)),
            headers: {
                "X-HTTP-Method": "MERGE",
            },
        }).then(function (d) { return mapper(d, props); }); };
    };
    return _SharePointQueryableInstance;
}(_SharePointQueryable));

var SharePointQueryableInstance = spInvokableFactory(_SharePointQueryableInstance);
/**
 * Adds the a delete method to the tagged class taking no parameters and calling spPostDelete
 */
function deleteable(t) {
    return function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPostDelete"])(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(this, t + ".delete"));
    };
}
function deleteableWithETag(t) {
    return function (eTag) {
        if (eTag === void 0) { eTag = "*"; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPostDeleteETag"])(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(this, t + ".delete"), {}, eTag);
    };
}
//# sourceMappingURL=sharepointqueryable.js.map

/***/ }),

/***/ "EVeO":
/*!********************************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/objectToSPKeyValueCollection.js ***!
  \********************************************************************/
/*! exports provided: objectToSPKeyValueCollection */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "objectToSPKeyValueCollection", function() { return objectToSPKeyValueCollection; });
/* harmony import */ var _metadata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./metadata */ "z067");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");


/**
 * Creates an object representing a SharePoint Collection(SP.KeyValue)
 *
 * @param obj The plain object defining the properties
 */
function objectToSPKeyValueCollection(obj) {
    return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_metadata__WEBPACK_IMPORTED_MODULE_0__["metadata"])("Collection(SP.KeyValue)"), {
        results: Object.keys(obj).map(function (key) { return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_metadata__WEBPACK_IMPORTED_MODULE_0__["metadata"])("SP.KeyValue"), {
            Key: key,
            Value: Reflect.get(obj, key),
            ValueType: "Edm.String",
        }); }),
    });
}
//# sourceMappingURL=objectToSPKeyValueCollection.js.map

/***/ }),

/***/ "EjWy":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/site-users/web.js ***!
  \************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "y+KB");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../operations */ "UK2s");






Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "siteUsers", _types__WEBPACK_IMPORTED_MODULE_3__["SiteUsers"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"], "currentUser", _types__WEBPACK_IMPORTED_MODULE_3__["SiteUser"], "currentuser");
_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"].prototype.ensureUser = function (logonName) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var data;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(_webs_types__WEBPACK_IMPORTED_MODULE_2__["Web"], "ensureuser"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ logonName: logonName }))];
                case 1:
                    data = _a.sent();
                    return [2 /*return*/, {
                            data: data,
                            user: Object(_types__WEBPACK_IMPORTED_MODULE_3__["SiteUser"])(Object(_odata__WEBPACK_IMPORTED_MODULE_4__["odataUrlFrom"])(data)),
                        }];
            }
        });
    });
};
_webs_types__WEBPACK_IMPORTED_MODULE_2__["_Web"].prototype.getUserById = function (id) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_3__["SiteUser"])(this, "getUserById(" + id + ")");
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "EyfK":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/hubsites/site.js ***!
  \***********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sites_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sites/types */ "ilAS");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../operations */ "UK2s");



_sites_types__WEBPACK_IMPORTED_MODULE_1__["_Site"].prototype.joinHubSite = function (siteId) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this.clone(_sites_types__WEBPACK_IMPORTED_MODULE_1__["Site"], "joinHubSite('" + siteId + "')"))];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
};
_sites_types__WEBPACK_IMPORTED_MODULE_1__["_Site"].prototype.registerHubSite = function () {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this.clone(_sites_types__WEBPACK_IMPORTED_MODULE_1__["Site"], "registerHubSite"))];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
};
_sites_types__WEBPACK_IMPORTED_MODULE_1__["_Site"].prototype.unRegisterHubSite = function () {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this.clone(_sites_types__WEBPACK_IMPORTED_MODULE_1__["Site"], "unRegisterHubSite"))];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
};
//# sourceMappingURL=site.js.map

/***/ }),

/***/ "F/m7":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/presets/all.js ***!
  \*********************************************/
/*! exports provided: spODataEntity, spODataEntityArray, sp, App, AppCatalog, Attachment, Attachments, ClientsidePageFromFile, CreateClientsidePage, CanvasColumn, CanvasSection, ClientsideText, ClientsideWebpart, ColumnControl, PromotedState, Comment, Comments, Replies, ContentType, ContentTypes, FieldLink, FieldLinks, Feature, Features, Fields, Field, AddFieldOptions, CalendarType, ChoiceFieldFormatType, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType, FieldTypes, FieldUserSelectionMode, UrlFieldFormatType, File, Files, CheckinType, MoveOperations, TemplateFileType, Version, Versions, Folder, Folders, Form, Forms, HubSite, HubSites, Item, Items, ItemVersion, ItemVersions, PagedItemCollection, List, Lists, ControlMode, RenderListDataOptions, Navigation, NavigationNode, NavigationNodes, NavigationService, Profiles, UrlZone, RegionalSettings, TimeZone, TimeZones, RelatedItemManager, SearchQueryBuilder, SearchResults, Search, Suggest, RoleAssignment, RoleAssignments, RoleDefinition, RoleDefinitions, PermissionKind, SPSharedObjectType, SharingDomainRestrictionMode, SharingLinkKind, SharingOperationStatusCode, SharingRole, RoleType, SiteDesigns, SiteGroup, SiteGroups, SiteScripts, SiteUser, SiteUsers, Site, MySocial, Social, SocialActorType, SocialActorTypes, SocialFollowResult, SocialStatusCode, Utilities, Subscription, Subscriptions, UserCustomAction, UserCustomActions, UserCustomActionRegistrationType, UserCustomActionScope, View, ViewFields, Views, ViewScope, WebPartsPersonalizationScope, WebPartDefinitions, WebPartDefinition, LimitedWebPartManager, Web, Webs, SharePointQueryableInstance, SharePointQueryableCollection, SharePointQueryable, spInvokableFactory, SPBatch, SPHttpClient, SPRest, toAbsoluteUrl, extractWebUrl, objectToSPKeyValueCollection, stripInvalidFileFolderChars, containsInvalidFileFolderChars, SortDirection, ReorderingRuleMatchType, QueryPropertyValueType, SearchBuiltInSourceId, defaultPath, spGet, spPost, spDelete, spPatch, spPostDelete, spPostDeleteETag, PrincipalType, PrincipalSource, PageType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "sp", function() { return sp; });
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _appcatalog__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../appcatalog */ "0FKx");
/* harmony import */ var _attachments__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../attachments */ "HLtT");
/* harmony import */ var _clientside_pages__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../clientside-pages */ "36LU");
/* harmony import */ var _column_defaults__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../column-defaults */ "pGaT");
/* harmony import */ var _comments__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../comments */ "KvhF");
/* harmony import */ var _content_types__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../content-types */ "dGPD");
/* harmony import */ var _features__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../features */ "CbhW");
/* harmony import */ var _fields__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../fields */ "wmuB");
/* harmony import */ var _files__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../files */ "0YVK");
/* harmony import */ var _folders__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../folders */ "0MvI");
/* harmony import */ var _forms__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../forms */ "GnYu");
/* harmony import */ var _hubsites__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../hubsites */ "Sdn1");
/* harmony import */ var _items__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ../items */ "lYrR");
/* harmony import */ var _lists__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! ../lists */ "J7sA");
/* harmony import */ var _navigation__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! ../navigation */ "ZdGM");
/* harmony import */ var _profiles__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! ../profiles */ "eKJE");
/* harmony import */ var _regional_settings__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(/*! ../regional-settings */ "pNrY");
/* harmony import */ var _related_items__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(/*! ../related-items */ "TUji");
/* harmony import */ var _search__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(/*! ../search */ "D2kI");
/* harmony import */ var _security__WEBPACK_IMPORTED_MODULE_20__ = __webpack_require__(/*! ../security */ "ah5w");
/* harmony import */ var _sharing__WEBPACK_IMPORTED_MODULE_21__ = __webpack_require__(/*! ../sharing */ "ikbv");
/* harmony import */ var _site_designs__WEBPACK_IMPORTED_MODULE_22__ = __webpack_require__(/*! ../site-designs */ "qoFj");
/* harmony import */ var _site_groups__WEBPACK_IMPORTED_MODULE_23__ = __webpack_require__(/*! ../site-groups */ "6l60");
/* harmony import */ var _site_scripts__WEBPACK_IMPORTED_MODULE_24__ = __webpack_require__(/*! ../site-scripts */ "UuEs");
/* harmony import */ var _site_users__WEBPACK_IMPORTED_MODULE_25__ = __webpack_require__(/*! ../site-users */ "vX5r");
/* harmony import */ var _sites__WEBPACK_IMPORTED_MODULE_26__ = __webpack_require__(/*! ../sites */ "1jZ6");
/* harmony import */ var _social__WEBPACK_IMPORTED_MODULE_27__ = __webpack_require__(/*! ../social */ "E/FW");
/* harmony import */ var _sputilities__WEBPACK_IMPORTED_MODULE_28__ = __webpack_require__(/*! ../sputilities */ "Ku5p");
/* harmony import */ var _subscriptions__WEBPACK_IMPORTED_MODULE_29__ = __webpack_require__(/*! ../subscriptions */ "Al5K");
/* harmony import */ var _user_custom_actions__WEBPACK_IMPORTED_MODULE_30__ = __webpack_require__(/*! ../user-custom-actions */ "HkJU");
/* harmony import */ var _views__WEBPACK_IMPORTED_MODULE_31__ = __webpack_require__(/*! ../views */ "dYBM");
/* harmony import */ var _webparts__WEBPACK_IMPORTED_MODULE_32__ = __webpack_require__(/*! ../webparts */ "9Ar0");
/* harmony import */ var _webs__WEBPACK_IMPORTED_MODULE_33__ = __webpack_require__(/*! ../webs */ "6k7F");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "App", function() { return _appcatalog__WEBPACK_IMPORTED_MODULE_1__["App"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "AppCatalog", function() { return _appcatalog__WEBPACK_IMPORTED_MODULE_1__["AppCatalog"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Attachment", function() { return _attachments__WEBPACK_IMPORTED_MODULE_2__["Attachment"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Attachments", function() { return _attachments__WEBPACK_IMPORTED_MODULE_2__["Attachments"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ClientsidePageFromFile", function() { return _clientside_pages__WEBPACK_IMPORTED_MODULE_3__["ClientsidePageFromFile"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CreateClientsidePage", function() { return _clientside_pages__WEBPACK_IMPORTED_MODULE_3__["CreateClientsidePage"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CanvasColumn", function() { return _clientside_pages__WEBPACK_IMPORTED_MODULE_3__["CanvasColumn"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CanvasSection", function() { return _clientside_pages__WEBPACK_IMPORTED_MODULE_3__["CanvasSection"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ClientsideText", function() { return _clientside_pages__WEBPACK_IMPORTED_MODULE_3__["ClientsideText"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ClientsideWebpart", function() { return _clientside_pages__WEBPACK_IMPORTED_MODULE_3__["ClientsideWebpart"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ColumnControl", function() { return _clientside_pages__WEBPACK_IMPORTED_MODULE_3__["ColumnControl"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PromotedState", function() { return _clientside_pages__WEBPACK_IMPORTED_MODULE_3__["PromotedState"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Comment", function() { return _comments__WEBPACK_IMPORTED_MODULE_5__["Comment"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Comments", function() { return _comments__WEBPACK_IMPORTED_MODULE_5__["Comments"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Replies", function() { return _comments__WEBPACK_IMPORTED_MODULE_5__["Replies"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ContentType", function() { return _content_types__WEBPACK_IMPORTED_MODULE_6__["ContentType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ContentTypes", function() { return _content_types__WEBPACK_IMPORTED_MODULE_6__["ContentTypes"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FieldLink", function() { return _content_types__WEBPACK_IMPORTED_MODULE_6__["FieldLink"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FieldLinks", function() { return _content_types__WEBPACK_IMPORTED_MODULE_6__["FieldLinks"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Feature", function() { return _features__WEBPACK_IMPORTED_MODULE_7__["Feature"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Features", function() { return _features__WEBPACK_IMPORTED_MODULE_7__["Features"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Fields", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["Fields"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Field", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["Field"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "AddFieldOptions", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["AddFieldOptions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CalendarType", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["CalendarType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ChoiceFieldFormatType", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["ChoiceFieldFormatType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "DateTimeFieldFormatType", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["DateTimeFieldFormatType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "DateTimeFieldFriendlyFormatType", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["DateTimeFieldFriendlyFormatType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FieldTypes", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["FieldTypes"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FieldUserSelectionMode", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["FieldUserSelectionMode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UrlFieldFormatType", function() { return _fields__WEBPACK_IMPORTED_MODULE_8__["UrlFieldFormatType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "File", function() { return _files__WEBPACK_IMPORTED_MODULE_9__["File"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Files", function() { return _files__WEBPACK_IMPORTED_MODULE_9__["Files"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CheckinType", function() { return _files__WEBPACK_IMPORTED_MODULE_9__["CheckinType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "MoveOperations", function() { return _files__WEBPACK_IMPORTED_MODULE_9__["MoveOperations"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "TemplateFileType", function() { return _files__WEBPACK_IMPORTED_MODULE_9__["TemplateFileType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Version", function() { return _files__WEBPACK_IMPORTED_MODULE_9__["Version"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Versions", function() { return _files__WEBPACK_IMPORTED_MODULE_9__["Versions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Folder", function() { return _folders__WEBPACK_IMPORTED_MODULE_10__["Folder"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Folders", function() { return _folders__WEBPACK_IMPORTED_MODULE_10__["Folders"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Form", function() { return _forms__WEBPACK_IMPORTED_MODULE_11__["Form"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Forms", function() { return _forms__WEBPACK_IMPORTED_MODULE_11__["Forms"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "HubSite", function() { return _hubsites__WEBPACK_IMPORTED_MODULE_12__["HubSite"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "HubSites", function() { return _hubsites__WEBPACK_IMPORTED_MODULE_12__["HubSites"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Item", function() { return _items__WEBPACK_IMPORTED_MODULE_13__["Item"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Items", function() { return _items__WEBPACK_IMPORTED_MODULE_13__["Items"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ItemVersion", function() { return _items__WEBPACK_IMPORTED_MODULE_13__["ItemVersion"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ItemVersions", function() { return _items__WEBPACK_IMPORTED_MODULE_13__["ItemVersions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PagedItemCollection", function() { return _items__WEBPACK_IMPORTED_MODULE_13__["PagedItemCollection"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "List", function() { return _lists__WEBPACK_IMPORTED_MODULE_14__["List"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Lists", function() { return _lists__WEBPACK_IMPORTED_MODULE_14__["Lists"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ControlMode", function() { return _lists__WEBPACK_IMPORTED_MODULE_14__["ControlMode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RenderListDataOptions", function() { return _lists__WEBPACK_IMPORTED_MODULE_14__["RenderListDataOptions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Navigation", function() { return _navigation__WEBPACK_IMPORTED_MODULE_15__["Navigation"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "NavigationNode", function() { return _navigation__WEBPACK_IMPORTED_MODULE_15__["NavigationNode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "NavigationNodes", function() { return _navigation__WEBPACK_IMPORTED_MODULE_15__["NavigationNodes"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "NavigationService", function() { return _navigation__WEBPACK_IMPORTED_MODULE_15__["NavigationService"]; });

/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_34__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spODataEntity", function() { return _odata__WEBPACK_IMPORTED_MODULE_34__["spODataEntity"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spODataEntityArray", function() { return _odata__WEBPACK_IMPORTED_MODULE_34__["spODataEntityArray"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Profiles", function() { return _profiles__WEBPACK_IMPORTED_MODULE_16__["Profiles"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UrlZone", function() { return _profiles__WEBPACK_IMPORTED_MODULE_16__["UrlZone"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RegionalSettings", function() { return _regional_settings__WEBPACK_IMPORTED_MODULE_17__["RegionalSettings"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "TimeZone", function() { return _regional_settings__WEBPACK_IMPORTED_MODULE_17__["TimeZone"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "TimeZones", function() { return _regional_settings__WEBPACK_IMPORTED_MODULE_17__["TimeZones"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RelatedItemManager", function() { return _related_items__WEBPACK_IMPORTED_MODULE_18__["RelatedItemManager"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SearchQueryBuilder", function() { return _search__WEBPACK_IMPORTED_MODULE_19__["SearchQueryBuilder"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SearchResults", function() { return _search__WEBPACK_IMPORTED_MODULE_19__["SearchResults"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Search", function() { return _search__WEBPACK_IMPORTED_MODULE_19__["Search"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Suggest", function() { return _search__WEBPACK_IMPORTED_MODULE_19__["Suggest"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SortDirection", function() { return _search__WEBPACK_IMPORTED_MODULE_19__["SortDirection"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ReorderingRuleMatchType", function() { return _search__WEBPACK_IMPORTED_MODULE_19__["ReorderingRuleMatchType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "QueryPropertyValueType", function() { return _search__WEBPACK_IMPORTED_MODULE_19__["QueryPropertyValueType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SearchBuiltInSourceId", function() { return _search__WEBPACK_IMPORTED_MODULE_19__["SearchBuiltInSourceId"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleAssignment", function() { return _security__WEBPACK_IMPORTED_MODULE_20__["RoleAssignment"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleAssignments", function() { return _security__WEBPACK_IMPORTED_MODULE_20__["RoleAssignments"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleDefinition", function() { return _security__WEBPACK_IMPORTED_MODULE_20__["RoleDefinition"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleDefinitions", function() { return _security__WEBPACK_IMPORTED_MODULE_20__["RoleDefinitions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PermissionKind", function() { return _security__WEBPACK_IMPORTED_MODULE_20__["PermissionKind"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPSharedObjectType", function() { return _sharing__WEBPACK_IMPORTED_MODULE_21__["SPSharedObjectType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharingDomainRestrictionMode", function() { return _sharing__WEBPACK_IMPORTED_MODULE_21__["SharingDomainRestrictionMode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharingLinkKind", function() { return _sharing__WEBPACK_IMPORTED_MODULE_21__["SharingLinkKind"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharingOperationStatusCode", function() { return _sharing__WEBPACK_IMPORTED_MODULE_21__["SharingOperationStatusCode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharingRole", function() { return _sharing__WEBPACK_IMPORTED_MODULE_21__["SharingRole"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleType", function() { return _sharing__WEBPACK_IMPORTED_MODULE_21__["RoleType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteDesigns", function() { return _site_designs__WEBPACK_IMPORTED_MODULE_22__["SiteDesigns"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteGroup", function() { return _site_groups__WEBPACK_IMPORTED_MODULE_23__["SiteGroup"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteGroups", function() { return _site_groups__WEBPACK_IMPORTED_MODULE_23__["SiteGroups"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteScripts", function() { return _site_scripts__WEBPACK_IMPORTED_MODULE_24__["SiteScripts"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteUser", function() { return _site_users__WEBPACK_IMPORTED_MODULE_25__["SiteUser"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteUsers", function() { return _site_users__WEBPACK_IMPORTED_MODULE_25__["SiteUsers"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Site", function() { return _sites__WEBPACK_IMPORTED_MODULE_26__["Site"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "MySocial", function() { return _social__WEBPACK_IMPORTED_MODULE_27__["MySocial"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Social", function() { return _social__WEBPACK_IMPORTED_MODULE_27__["Social"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SocialActorType", function() { return _social__WEBPACK_IMPORTED_MODULE_27__["SocialActorType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SocialActorTypes", function() { return _social__WEBPACK_IMPORTED_MODULE_27__["SocialActorTypes"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SocialFollowResult", function() { return _social__WEBPACK_IMPORTED_MODULE_27__["SocialFollowResult"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SocialStatusCode", function() { return _social__WEBPACK_IMPORTED_MODULE_27__["SocialStatusCode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Utilities", function() { return _sputilities__WEBPACK_IMPORTED_MODULE_28__["Utilities"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Subscription", function() { return _subscriptions__WEBPACK_IMPORTED_MODULE_29__["Subscription"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Subscriptions", function() { return _subscriptions__WEBPACK_IMPORTED_MODULE_29__["Subscriptions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UserCustomAction", function() { return _user_custom_actions__WEBPACK_IMPORTED_MODULE_30__["UserCustomAction"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UserCustomActions", function() { return _user_custom_actions__WEBPACK_IMPORTED_MODULE_30__["UserCustomActions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UserCustomActionRegistrationType", function() { return _user_custom_actions__WEBPACK_IMPORTED_MODULE_30__["UserCustomActionRegistrationType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UserCustomActionScope", function() { return _user_custom_actions__WEBPACK_IMPORTED_MODULE_30__["UserCustomActionScope"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "View", function() { return _views__WEBPACK_IMPORTED_MODULE_31__["View"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ViewFields", function() { return _views__WEBPACK_IMPORTED_MODULE_31__["ViewFields"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Views", function() { return _views__WEBPACK_IMPORTED_MODULE_31__["Views"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ViewScope", function() { return _views__WEBPACK_IMPORTED_MODULE_31__["ViewScope"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "WebPartsPersonalizationScope", function() { return _webparts__WEBPACK_IMPORTED_MODULE_32__["WebPartsPersonalizationScope"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "WebPartDefinitions", function() { return _webparts__WEBPACK_IMPORTED_MODULE_32__["WebPartDefinitions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "WebPartDefinition", function() { return _webparts__WEBPACK_IMPORTED_MODULE_32__["WebPartDefinition"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "LimitedWebPartManager", function() { return _webparts__WEBPACK_IMPORTED_MODULE_32__["LimitedWebPartManager"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Web", function() { return _webs__WEBPACK_IMPORTED_MODULE_33__["Web"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Webs", function() { return _webs__WEBPACK_IMPORTED_MODULE_33__["Webs"]; });

/* harmony import */ var _index__WEBPACK_IMPORTED_MODULE_35__ = __webpack_require__(/*! ../index */ "UKGb");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryableInstance", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["SharePointQueryableInstance"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryableCollection", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["SharePointQueryableCollection"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryable", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["SharePointQueryable"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spInvokableFactory", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["spInvokableFactory"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPBatch", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["SPBatch"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPHttpClient", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["SPHttpClient"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPRest", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["SPRest"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "toAbsoluteUrl", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["toAbsoluteUrl"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "extractWebUrl", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["extractWebUrl"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "objectToSPKeyValueCollection", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["objectToSPKeyValueCollection"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "stripInvalidFileFolderChars", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["stripInvalidFileFolderChars"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "containsInvalidFileFolderChars", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["containsInvalidFileFolderChars"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "defaultPath", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["defaultPath"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spGet", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["spGet"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spPost", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["spPost"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spDelete", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["spDelete"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spPatch", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["spPatch"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spPostDelete", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["spPostDelete"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spPostDeleteETag", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["spPostDeleteETag"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PrincipalType", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["PrincipalType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PrincipalSource", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["PrincipalSource"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PageType", function() { return _index__WEBPACK_IMPORTED_MODULE_35__["PageType"]; });





































































var sp = new _rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"]();
//# sourceMappingURL=all.js.map

/***/ }),

/***/ "Fx3T":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/attachments/item.js ***!
  \**************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "Ar68");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"], "attachmentFiles", _types__WEBPACK_IMPORTED_MODULE_2__["Attachments"]);
//# sourceMappingURL=item.js.map

/***/ }),

/***/ "G2Ar":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/toabsoluteurl.js ***!
  \*****************************************************/
/*! exports provided: toAbsoluteUrl */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "toAbsoluteUrl", function() { return toAbsoluteUrl; });
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _splibconfig__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../splibconfig */ "0p/9");


/**
 * Ensures that a given url is absolute for the current web based on context
 *
 * @param candidateUrl The url to make absolute
 *
 */
function toAbsoluteUrl(candidateUrl) {
    return new Promise(function (resolve) {
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["isUrlAbsolute"])(candidateUrl)) {
            // if we are already absolute, then just return the url
            return resolve(candidateUrl);
        }
        if (_splibconfig__WEBPACK_IMPORTED_MODULE_1__["SPRuntimeConfig"].baseUrl !== null) {
            // base url specified either with baseUrl of spfxContext config property
            return resolve(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["combine"])(_splibconfig__WEBPACK_IMPORTED_MODULE_1__["SPRuntimeConfig"].baseUrl, candidateUrl));
        }
        if (_pnp_common__WEBPACK_IMPORTED_MODULE_0__["safeGlobal"]._spPageContextInfo !== undefined) {
            // operating in classic pages
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["hOP"])(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["safeGlobal"]._spPageContextInfo, "webAbsoluteUrl")) {
                return resolve(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["combine"])(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["safeGlobal"]._spPageContextInfo.webAbsoluteUrl, candidateUrl));
            }
            else if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["hOP"])(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["safeGlobal"]._spPageContextInfo, "webServerRelativeUrl")) {
                return resolve(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["combine"])(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["safeGlobal"]._spPageContextInfo.webServerRelativeUrl, candidateUrl));
            }
        }
        // does window.location exist and have a certain path part in it?
        if (_pnp_common__WEBPACK_IMPORTED_MODULE_0__["safeGlobal"].location !== undefined) {
            var baseUrl_1 = _pnp_common__WEBPACK_IMPORTED_MODULE_0__["safeGlobal"].location.toString().toLowerCase();
            ["/_layouts/", "/siteassets/"].forEach(function (s) {
                var index = baseUrl_1.indexOf(s);
                if (index > 0) {
                    return resolve(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["combine"])(baseUrl_1.substr(0, index), candidateUrl));
                }
            });
        }
        return resolve(candidateUrl);
    });
}
//# sourceMappingURL=toabsoluteurl.js.map

/***/ }),

/***/ "GPet":
/*!*************************************************!*\
  !*** external "@microsoft/sp-application-base" ***!
  \*************************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_GPet__;

/***/ }),

/***/ "GnYu":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/forms/index.js ***!
  \*********************************************/
/*! exports provided: Form, Forms */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list */ "3fSn");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "zYJ9");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Form", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Form"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Forms", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Forms"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "HHhI":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/features/types.js ***!
  \************************************************/
/*! exports provided: _Features, Features, _Feature, Feature */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Features", function() { return _Features; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Features", function() { return Features; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Feature", function() { return _Feature; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Feature", function() { return Feature; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../telemetry */ "h3gW");






var _Features = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Features, _super);
    function _Features() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Adds (activates) the specified feature
     *
     * @param id The Id of the feature (GUID)
     * @param force If true the feature activation will be forced
     */
    _Features.prototype.add = function (id, force) {
        if (force === void 0) { force = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(this.clone(Features, "add"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({
                            featdefScope: 0,
                            featureId: id,
                            force: force,
                        }))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                feature: this.getById(id),
                            }];
                }
            });
        });
    };
    /**
     * Gets a feature from the collection with the specified guid
     *
     * @param id The Id of the feature (GUID)
     */
    _Features.prototype.getById = function (id) {
        var feature = Feature(this);
        feature.concat("('" + id + "')");
        return _telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"].configure(feature, "fes.getById");
    };
    /**
     * Removes (deactivates) a feature from the collection
     *
     * @param id The Id of the feature (GUID)
     * @param force If true the feature deactivation will be forced
     */
    _Features.prototype.remove = function (id, force) {
        if (force === void 0) { force = false; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(this.clone(Features, "remove"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({
            featureId: id,
            force: force,
        }));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("fes.add")
    ], _Features.prototype, "add", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("fes.remove")
    ], _Features.prototype, "remove", null);
    _Features = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("features")
    ], _Features);
    return _Features;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var Features = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_Features);
var _Feature = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Feature, _super);
    function _Feature() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Removes (deactivates) the feature
     *
     * @param force If true the feature deactivation will be forced
     */
    _Feature.prototype.deactivate = function (force) {
        if (force === void 0) { force = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var removeDependency, feature, promise;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        removeDependency = this.addBatchDependency();
                        return [4 /*yield*/, Feature(this).select("DefinitionId")()];
                    case 1:
                        feature = _a.sent();
                        promise = this.getParent(Features, this.parentUrl, "", this.batch).remove(feature.DefinitionId, force);
                        removeDependency();
                        return [2 /*return*/, promise];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("fe.deactivate")
    ], _Feature.prototype, "deactivate", null);
    return _Feature;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableInstance"]));

var Feature = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_Feature);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "HLtT":
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/attachments/index.js ***!
  \***************************************************/
/*! exports provided: Attachment, Attachments */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _item__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./item */ "Fx3T");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "Ar68");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Attachment", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Attachment"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Attachments", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Attachments"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "HkJU":
/*!***********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/index.js ***!
  \***********************************************************/
/*! exports provided: UserCustomAction, UserCustomActions, UserCustomActionRegistrationType, UserCustomActionScope */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list */ "VnYa");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./web */ "uMGh");
/* harmony import */ var _site__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./site */ "7pU6");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "1mTr");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UserCustomAction", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["UserCustomAction"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UserCustomActions", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["UserCustomActions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UserCustomActionRegistrationType", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["UserCustomActionRegistrationType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UserCustomActionScope", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["UserCustomActionScope"]; });





//# sourceMappingURL=index.js.map

/***/ }),

/***/ "J7sA":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/lists/index.js ***!
  \*********************************************/
/*! exports provided: List, Lists, ControlMode, RenderListDataOptions */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "Bwa7");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "hy0S");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "List", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["List"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Lists", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Lists"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ControlMode", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ControlMode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RenderListDataOptions", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["RenderListDataOptions"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "JFDc":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/files/folder.js ***!
  \**********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _folders_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../folders/types */ "4tXH");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "qfcF");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"], "files", _types__WEBPACK_IMPORTED_MODULE_2__["Files"]);
//# sourceMappingURL=folder.js.map

/***/ }),

/***/ "JT9m":
/*!*******************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/web.js ***!
  \*******************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "luMW");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./funcs */ "T2j4");




Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "regionalSettings", _types__WEBPACK_IMPORTED_MODULE_2__["RegionalSettings"]);
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.titleResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_3__["getValueForUICultureBinder"])("titleResource");
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.descriptionResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_3__["getValueForUICultureBinder"])("descriptionResource");
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "K1m8":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/item.js ***!
  \**********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "AvPO");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./funcs */ "At/s");



_items_types__WEBPACK_IMPORTED_MODULE_0__["_Item"].prototype.shareWith = function (loginNames, role, requireSignin, emailData) {
    if (role === void 0) { role = _types__WEBPACK_IMPORTED_MODULE_1__["SharingRole"].View; }
    if (requireSignin === void 0) { requireSignin = false; }
    return Object(_funcs__WEBPACK_IMPORTED_MODULE_2__["shareWith"])(this, loginNames, role, requireSignin, false, emailData);
};
_items_types__WEBPACK_IMPORTED_MODULE_0__["_Item"].prototype.getShareLink = _funcs__WEBPACK_IMPORTED_MODULE_2__["getShareLink"];
_items_types__WEBPACK_IMPORTED_MODULE_0__["_Item"].prototype.checkSharingPermissions = _funcs__WEBPACK_IMPORTED_MODULE_2__["checkPermissions"];
_items_types__WEBPACK_IMPORTED_MODULE_0__["_Item"].prototype.getSharingInformation = _funcs__WEBPACK_IMPORTED_MODULE_2__["getSharingInformation"];
_items_types__WEBPACK_IMPORTED_MODULE_0__["_Item"].prototype.getObjectSharingSettings = _funcs__WEBPACK_IMPORTED_MODULE_2__["getObjectSharingSettings"];
_items_types__WEBPACK_IMPORTED_MODULE_0__["_Item"].prototype.unshare = _funcs__WEBPACK_IMPORTED_MODULE_2__["unshareObject"];
_items_types__WEBPACK_IMPORTED_MODULE_0__["_Item"].prototype.deleteSharingLinkByKind = _funcs__WEBPACK_IMPORTED_MODULE_2__["deleteLinkByKind"];
_items_types__WEBPACK_IMPORTED_MODULE_0__["_Item"].prototype.unshareLink = _funcs__WEBPACK_IMPORTED_MODULE_2__["unshareLink"];
//# sourceMappingURL=item.js.map

/***/ }),

/***/ "KJ7d":
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/list.js ***!
  \********************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./funcs */ "T2j4");


_lists_types__WEBPACK_IMPORTED_MODULE_0__["_List"].prototype.titleResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_1__["getValueForUICultureBinder"])("titleResource");
_lists_types__WEBPACK_IMPORTED_MODULE_0__["_List"].prototype.descriptionResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_1__["getValueForUICultureBinder"])("descriptionResource");
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "KL1q":
/*!***********************************************************!*\
  !*** external "@microsoft/office-ui-fabric-react-bundle" ***!
  \***********************************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_KL1q__;

/***/ }),

/***/ "KYlJ":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/fields/types.js ***!
  \**********************************************/
/*! exports provided: _Fields, Fields, _Field, Field, FieldTypes, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType, AddFieldOptions, CalendarType, UrlFieldFormatType, FieldUserSelectionMode, ChoiceFieldFormatType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Fields", function() { return _Fields; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Fields", function() { return Fields; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Field", function() { return _Field; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Field", function() { return Field; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "FieldTypes", function() { return FieldTypes; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "DateTimeFieldFormatType", function() { return DateTimeFieldFormatType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "DateTimeFieldFriendlyFormatType", function() { return DateTimeFieldFriendlyFormatType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "AddFieldOptions", function() { return AddFieldOptions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CalendarType", function() { return CalendarType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "UrlFieldFormatType", function() { return UrlFieldFormatType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "FieldUserSelectionMode", function() { return FieldUserSelectionMode; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ChoiceFieldFormatType", function() { return ChoiceFieldFormatType; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../telemetry */ "h3gW");








var _Fields = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Fields, _super);
    function _Fields() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a field from the collection by id
     *
     * @param id The Id of the list
     */
    _Fields.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].configure(Field(this).concat("('" + id + "')"), "fs.getById");
    };
    /**
     * Gets a field from the collection by title
     *
     * @param title The case-sensitive title of the field
     */
    _Fields.prototype.getByTitle = function (title) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].configure(Field(this, "getByTitle('" + title + "')"), "fs.getByTitle");
    };
    /**
     * Gets a field from the collection by using internal name or title
     *
     * @param name The case-sensitive internal name or title of the field
     */
    _Fields.prototype.getByInternalNameOrTitle = function (name) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].configure(Field(this, "getByInternalNameOrTitle('" + name + "')"), "fs.getByInternalNameOrTitle");
    };
    /**
     * Creates a field based on the specified schema
     *
     * @param xml A string or XmlSchemaFieldCreationInformation instance descrbing the field to create
     */
    _Fields.prototype.createFieldAsXml = function (xml) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (typeof xml === "string") {
                            xml = { SchemaXml: xml };
                        }
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({
                            "parameters": Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])("SP.XmlSchemaFieldCreationInformation"), xml),
                        });
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Fields, "createfieldasxml"), postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                field: this.getById(data.Id),
                            }];
                }
            });
        });
    };
    /**
     * Adds a new field to the collection
     *
     * @param title The new field's title
     * @param fieldType The new field's type (ex: SP.FieldText)
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.add = function (title, fieldType, properties) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])(Object.assign(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])(fieldType), {
                            "Title": title,
                        }, properties));
                        if (!_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].isTagged(this)) {
                            _telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"].configure(this, "fs.add");
                        }
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Fields, null), postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                field: this.getById(data.Id),
                            }];
                }
            });
        });
    };
    /**
     * Adds a new SP.FieldText to the collection
     *
     * @param title The field title
     * @param maxLength The maximum number of characters allowed in the value of the field.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addText = function (title, maxLength, properties) {
        if (maxLength === void 0) { maxLength = 255; }
        var props = {
            FieldTypeKind: 2,
            MaxLength: maxLength,
        };
        return this.add(title, "SP.FieldText", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a new SP.FieldCalculated to the collection
     *
     * @param title The field title.
     * @param formula The formula for the field.
     * @param dateFormat The date and time format that is displayed in the field.
     * @param outputType Specifies the output format for the field. Represents a FieldType value.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addCalculated = function (title, formula, dateFormat, outputType, properties) {
        if (outputType === void 0) { outputType = FieldTypes.Text; }
        var props = {
            DateFormat: dateFormat,
            FieldTypeKind: 17,
            Formula: formula,
            OutputType: outputType,
        };
        return this.add(title, "SP.FieldCalculated", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a new SP.FieldDateTime to the collection
     *
     * @param title The field title
     * @param displayFormat The format of the date and time that is displayed in the field.
     * @param calendarType Specifies the calendar type of the field.
     * @param friendlyDisplayFormat The type of friendly display format that is used in the field.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addDateTime = function (title, displayFormat, calendarType, friendlyDisplayFormat, properties) {
        if (displayFormat === void 0) { displayFormat = DateTimeFieldFormatType.DateOnly; }
        if (calendarType === void 0) { calendarType = CalendarType.Gregorian; }
        if (friendlyDisplayFormat === void 0) { friendlyDisplayFormat = DateTimeFieldFriendlyFormatType.Unspecified; }
        var props = {
            DateTimeCalendarType: calendarType,
            DisplayFormat: displayFormat,
            FieldTypeKind: 4,
            FriendlyDisplayFormat: friendlyDisplayFormat,
        };
        return this.add(title, "SP.FieldDateTime", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a new SP.FieldNumber to the collection
     *
     * @param title The field title
     * @param minValue The field's minimum value
     * @param maxValue The field's maximum value
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addNumber = function (title, minValue, maxValue, properties) {
        var props = { FieldTypeKind: 9 };
        if (minValue !== undefined) {
            props = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])({ MinimumValue: minValue }, props);
        }
        if (maxValue !== undefined) {
            props = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])({ MaximumValue: maxValue }, props);
        }
        return this.add(title, "SP.FieldNumber", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a new SP.FieldCurrency to the collection
     *
     * @param title The field title
     * @param minValue The field's minimum value
     * @param maxValue The field's maximum value
     * @param currencyLocalId Specifies the language code identifier (LCID) used to format the value of the field
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addCurrency = function (title, minValue, maxValue, currencyLocalId, properties) {
        if (currencyLocalId === void 0) { currencyLocalId = 1033; }
        var props = {
            CurrencyLocaleId: currencyLocalId,
            FieldTypeKind: 10,
        };
        if (minValue !== undefined) {
            props = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])({ MinimumValue: minValue }, props);
        }
        if (maxValue !== undefined) {
            props = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])({ MaximumValue: maxValue }, props);
        }
        return this.add(title, "SP.FieldCurrency", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a new SP.FieldMultiLineText to the collection
     *
     * @param title The field title
     * @param numberOfLines Specifies the number of lines of text to display for the field.
     * @param richText Specifies whether the field supports rich formatting.
     * @param restrictedMode Specifies whether the field supports a subset of rich formatting.
     * @param appendOnly Specifies whether all changes to the value of the field are displayed in list forms.
     * @param allowHyperlink Specifies whether a hyperlink is allowed as a value of the field.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     *
     */
    _Fields.prototype.addMultilineText = function (title, numberOfLines, richText, restrictedMode, appendOnly, allowHyperlink, properties) {
        if (numberOfLines === void 0) { numberOfLines = 6; }
        if (richText === void 0) { richText = true; }
        if (restrictedMode === void 0) { restrictedMode = false; }
        if (appendOnly === void 0) { appendOnly = false; }
        if (allowHyperlink === void 0) { allowHyperlink = true; }
        var props = {
            AllowHyperlink: allowHyperlink,
            AppendOnly: appendOnly,
            FieldTypeKind: 3,
            NumberOfLines: numberOfLines,
            RestrictedMode: restrictedMode,
            RichText: richText,
        };
        return this.add(title, "SP.FieldMultiLineText", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a new SP.FieldUrl to the collection
     *
     * @param title The field title
     */
    _Fields.prototype.addUrl = function (title, displayFormat, properties) {
        if (displayFormat === void 0) { displayFormat = UrlFieldFormatType.Hyperlink; }
        var props = {
            DisplayFormat: displayFormat,
            FieldTypeKind: 11,
        };
        return this.add(title, "SP.FieldUrl", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /** Adds a user field to the colleciton
    *
    * @param title The new field's title
    * @param selectionMode The selection mode of the field
    * @param selectionGroup Value that specifies the identifier of the SharePoint group whose members can be selected as values of the field
    * @param properties
    */
    _Fields.prototype.addUser = function (title, selectionMode, properties) {
        var props = {
            FieldTypeKind: 20,
            SelectionMode: selectionMode,
        };
        return this.add(title, "SP.FieldUser", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a SP.FieldLookup to the collection
     *
     * @param title The new field's title
     * @param lookupListId The guid id of the list where the source of the lookup is found
     * @param lookupFieldName The internal name of the field in the source list
     * @param properties Set of additional properties to set on the new field
     */
    _Fields.prototype.addLookup = function (title, lookupListId, lookupFieldName, properties) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var props, postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        props = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])({
                            FieldTypeKind: 7,
                            LookupFieldName: lookupFieldName,
                            LookupListId: lookupListId,
                            Title: title,
                        }, properties);
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({
                            "parameters": Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])("SP.FieldCreationInformation"), props),
                        });
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Fields, "addfield"), postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                field: this.getById(data.Id),
                            }];
                }
            });
        });
    };
    /**
     * Adds a new SP.FieldChoice to the collection
     *
     * @param title The field title.
     * @param choices The choices for the field.
     * @param format The display format of the available options for the field.
     * @param fillIn Specifies whether the field allows fill-in values.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addChoice = function (title, choices, format, fillIn, properties) {
        if (format === void 0) { format = ChoiceFieldFormatType.Dropdown; }
        var props = {
            Choices: {
                results: choices,
            },
            EditFormat: format,
            FieldTypeKind: 6,
            FillInChoice: fillIn,
        };
        return this.add(title, "SP.FieldChoice", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a new SP.FieldMultiChoice to the collection
     *
     * @param title The field title.
     * @param choices The choices for the field.
     * @param fillIn Specifies whether the field allows fill-in values.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addMultiChoice = function (title, choices, fillIn, properties) {
        var props = {
            Choices: {
                results: choices,
            },
            FieldTypeKind: 15,
            FillInChoice: fillIn,
        };
        return this.add(title, "SP.FieldMultiChoice", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
     * Adds a new SP.FieldBoolean to the collection
     *
     * @param title The field title.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addBoolean = function (title, properties) {
        var props = {
            FieldTypeKind: 8,
        };
        return this.add(title, "SP.Field", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    /**
    * Creates a secondary (dependent) lookup field, based on the Id of the primary lookup field.
    *
    * @param displayName The display name of the new field.
    * @param primaryLookupFieldId The guid of the primary Lookup Field.
    * @param showField Which field to show from the lookup list.
    */
    _Fields.prototype.addDependentLookupField = function (displayName, primaryLookupFieldId, showField) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var path, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        path = "adddependentlookupfield(displayName='" + displayName + "', primarylookupfieldid='" + primaryLookupFieldId + "', showfield='" + showField + "')";
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Fields, path))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                field: this.getById(data.Id),
                            }];
                }
            });
        });
    };
    /**
     * Adds a new SP.FieldLocation to the collection
     *
     * @param title The field title.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    _Fields.prototype.addLocation = function (title, properties) {
        var props = { FieldTypeKind: 33 };
        return this.add(title, "SP.FieldLocation", Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(props, properties));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.createFieldAsXml")
    ], _Fields.prototype, "createFieldAsXml", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addText")
    ], _Fields.prototype, "addText", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addCalculated")
    ], _Fields.prototype, "addCalculated", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addDateTime")
    ], _Fields.prototype, "addDateTime", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addNumber")
    ], _Fields.prototype, "addNumber", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addCurrency")
    ], _Fields.prototype, "addCurrency", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addMultilineText")
    ], _Fields.prototype, "addMultilineText", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addUrl")
    ], _Fields.prototype, "addUrl", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addUser")
    ], _Fields.prototype, "addUser", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addLookup")
    ], _Fields.prototype, "addLookup", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addChoice")
    ], _Fields.prototype, "addChoice", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addMultiChoice")
    ], _Fields.prototype, "addMultiChoice", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addBoolean")
    ], _Fields.prototype, "addBoolean", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addDependentLookupField")
    ], _Fields.prototype, "addDependentLookupField", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("fs.addLocation")
    ], _Fields.prototype, "addLocation", null);
    _Fields = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_5__["defaultPath"])("fields")
    ], _Fields);
    return _Fields;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var Fields = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Fields);
var _Field = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Field, _super);
    function _Field() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["deleteable"])("f");
        return _this;
    }
    /**
     * Updates this field instance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the list
     * @param fieldType The type value such as SP.FieldLookup. Optional, looked up from the field if not provided
     */
    _Field.prototype.update = function (properties, fieldType) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var info, req, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(typeof fieldType === "undefined" || fieldType === null)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.select("FieldTypeKind")()];
                    case 1:
                        info = _a.sent();
                        fieldType = "SP.Field" + FieldTypes[info.FieldTypeKind];
                        _a.label = 2;
                    case 2:
                        req = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])(fieldType), properties), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["headers"])({ "X-HTTP-Method": "MERGE" }));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this, req)];
                    case 3:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                field: this,
                            }];
                }
            });
        });
    };
    /**
     * Sets the value of the ShowInDisplayForm property for this field.
     */
    _Field.prototype.setShowInDisplayForm = function (show) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Field, "setshowindisplayform(" + show + ")"));
    };
    /**
     * Sets the value of the ShowInEditForm property for this field.
     */
    _Field.prototype.setShowInEditForm = function (show) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Field, "setshowineditform(" + show + ")"));
    };
    /**
     * Sets the value of the ShowInNewForm property for this field.
     */
    _Field.prototype.setShowInNewForm = function (show) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_6__["spPost"])(this.clone(Field, "setshowinnewform(" + show + ")"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("f.update")
    ], _Field.prototype, "update", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("f.setShowInDisplayForm")
    ], _Field.prototype, "setShowInDisplayForm", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("f.setShowInEditForm")
    ], _Field.prototype, "setShowInEditForm", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("f.setShowInNewForm")
    ], _Field.prototype, "setShowInNewForm", null);
    return _Field;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Field = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Field);
/**
 * Specifies the type of the field.
 */
var FieldTypes;
(function (FieldTypes) {
    FieldTypes[FieldTypes["Invalid"] = 0] = "Invalid";
    FieldTypes[FieldTypes["Integer"] = 1] = "Integer";
    FieldTypes[FieldTypes["Text"] = 2] = "Text";
    FieldTypes[FieldTypes["Note"] = 3] = "Note";
    FieldTypes[FieldTypes["DateTime"] = 4] = "DateTime";
    FieldTypes[FieldTypes["Counter"] = 5] = "Counter";
    FieldTypes[FieldTypes["Choice"] = 6] = "Choice";
    FieldTypes[FieldTypes["Lookup"] = 7] = "Lookup";
    FieldTypes[FieldTypes["Boolean"] = 8] = "Boolean";
    FieldTypes[FieldTypes["Number"] = 9] = "Number";
    FieldTypes[FieldTypes["Currency"] = 10] = "Currency";
    FieldTypes[FieldTypes["URL"] = 11] = "URL";
    FieldTypes[FieldTypes["Computed"] = 12] = "Computed";
    FieldTypes[FieldTypes["Threading"] = 13] = "Threading";
    FieldTypes[FieldTypes["Guid"] = 14] = "Guid";
    FieldTypes[FieldTypes["MultiChoice"] = 15] = "MultiChoice";
    FieldTypes[FieldTypes["GridChoice"] = 16] = "GridChoice";
    FieldTypes[FieldTypes["Calculated"] = 17] = "Calculated";
    FieldTypes[FieldTypes["File"] = 18] = "File";
    FieldTypes[FieldTypes["Attachments"] = 19] = "Attachments";
    FieldTypes[FieldTypes["User"] = 20] = "User";
    FieldTypes[FieldTypes["Recurrence"] = 21] = "Recurrence";
    FieldTypes[FieldTypes["CrossProjectLink"] = 22] = "CrossProjectLink";
    FieldTypes[FieldTypes["ModStat"] = 23] = "ModStat";
    FieldTypes[FieldTypes["Error"] = 24] = "Error";
    FieldTypes[FieldTypes["ContentTypeId"] = 25] = "ContentTypeId";
    FieldTypes[FieldTypes["PageSeparator"] = 26] = "PageSeparator";
    FieldTypes[FieldTypes["ThreadIndex"] = 27] = "ThreadIndex";
    FieldTypes[FieldTypes["WorkflowStatus"] = 28] = "WorkflowStatus";
    FieldTypes[FieldTypes["AllDayEvent"] = 29] = "AllDayEvent";
    FieldTypes[FieldTypes["WorkflowEventType"] = 30] = "WorkflowEventType";
})(FieldTypes || (FieldTypes = {}));
var DateTimeFieldFormatType;
(function (DateTimeFieldFormatType) {
    DateTimeFieldFormatType[DateTimeFieldFormatType["DateOnly"] = 0] = "DateOnly";
    DateTimeFieldFormatType[DateTimeFieldFormatType["DateTime"] = 1] = "DateTime";
})(DateTimeFieldFormatType || (DateTimeFieldFormatType = {}));
var DateTimeFieldFriendlyFormatType;
(function (DateTimeFieldFriendlyFormatType) {
    DateTimeFieldFriendlyFormatType[DateTimeFieldFriendlyFormatType["Unspecified"] = 0] = "Unspecified";
    DateTimeFieldFriendlyFormatType[DateTimeFieldFriendlyFormatType["Disabled"] = 1] = "Disabled";
    DateTimeFieldFriendlyFormatType[DateTimeFieldFriendlyFormatType["Relative"] = 2] = "Relative";
})(DateTimeFieldFriendlyFormatType || (DateTimeFieldFriendlyFormatType = {}));
/**
 * Specifies the control settings while adding a field.
 */
var AddFieldOptions;
(function (AddFieldOptions) {
    /**
     *  Specify that a new field added to the list must also be added to the default content type in the site collection
     */
    AddFieldOptions[AddFieldOptions["DefaultValue"] = 0] = "DefaultValue";
    /**
     * Specify that a new field added to the list must also be added to the default content type in the site collection.
     */
    AddFieldOptions[AddFieldOptions["AddToDefaultContentType"] = 1] = "AddToDefaultContentType";
    /**
     * Specify that a new field must not be added to any other content type
     */
    AddFieldOptions[AddFieldOptions["AddToNoContentType"] = 2] = "AddToNoContentType";
    /**
     *  Specify that a new field that is added to the specified list must also be added to all content types in the site collection
     */
    AddFieldOptions[AddFieldOptions["AddToAllContentTypes"] = 4] = "AddToAllContentTypes";
    /**
     * Specify adding an internal field name hint for the purpose of avoiding possible database locking or field renaming operations
     */
    AddFieldOptions[AddFieldOptions["AddFieldInternalNameHint"] = 8] = "AddFieldInternalNameHint";
    /**
     * Specify that a new field that is added to the specified list must also be added to the default list view
     */
    AddFieldOptions[AddFieldOptions["AddFieldToDefaultView"] = 16] = "AddFieldToDefaultView";
    /**
     * Specify to confirm that no other field has the same display name
     */
    AddFieldOptions[AddFieldOptions["AddFieldCheckDisplayName"] = 32] = "AddFieldCheckDisplayName";
})(AddFieldOptions || (AddFieldOptions = {}));
var CalendarType;
(function (CalendarType) {
    CalendarType[CalendarType["Gregorian"] = 1] = "Gregorian";
    CalendarType[CalendarType["Japan"] = 3] = "Japan";
    CalendarType[CalendarType["Taiwan"] = 4] = "Taiwan";
    CalendarType[CalendarType["Korea"] = 5] = "Korea";
    CalendarType[CalendarType["Hijri"] = 6] = "Hijri";
    CalendarType[CalendarType["Thai"] = 7] = "Thai";
    CalendarType[CalendarType["Hebrew"] = 8] = "Hebrew";
    CalendarType[CalendarType["GregorianMEFrench"] = 9] = "GregorianMEFrench";
    CalendarType[CalendarType["GregorianArabic"] = 10] = "GregorianArabic";
    CalendarType[CalendarType["GregorianXLITEnglish"] = 11] = "GregorianXLITEnglish";
    CalendarType[CalendarType["GregorianXLITFrench"] = 12] = "GregorianXLITFrench";
    CalendarType[CalendarType["KoreaJapanLunar"] = 14] = "KoreaJapanLunar";
    CalendarType[CalendarType["ChineseLunar"] = 15] = "ChineseLunar";
    CalendarType[CalendarType["SakaEra"] = 16] = "SakaEra";
    CalendarType[CalendarType["UmAlQura"] = 23] = "UmAlQura";
})(CalendarType || (CalendarType = {}));
var UrlFieldFormatType;
(function (UrlFieldFormatType) {
    UrlFieldFormatType[UrlFieldFormatType["Hyperlink"] = 0] = "Hyperlink";
    UrlFieldFormatType[UrlFieldFormatType["Image"] = 1] = "Image";
})(UrlFieldFormatType || (UrlFieldFormatType = {}));
var FieldUserSelectionMode;
(function (FieldUserSelectionMode) {
    FieldUserSelectionMode[FieldUserSelectionMode["PeopleAndGroups"] = 1] = "PeopleAndGroups";
    FieldUserSelectionMode[FieldUserSelectionMode["PeopleOnly"] = 0] = "PeopleOnly";
})(FieldUserSelectionMode || (FieldUserSelectionMode = {}));
var ChoiceFieldFormatType;
(function (ChoiceFieldFormatType) {
    ChoiceFieldFormatType[ChoiceFieldFormatType["Dropdown"] = 0] = "Dropdown";
    ChoiceFieldFormatType[ChoiceFieldFormatType["RadioButtons"] = 1] = "RadioButtons";
})(ChoiceFieldFormatType || (ChoiceFieldFormatType = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "Ka72":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/security/item.js ***!
  \***********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "1dmT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./funcs */ "vXSC");





Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"], "roleAssignments", _types__WEBPACK_IMPORTED_MODULE_2__["RoleAssignments"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"], "firstUniqueAncestorSecurableObject", _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"]);
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.getUserEffectivePermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["getUserEffectivePermissions"];
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.getCurrentUserEffectivePermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["getCurrentUserEffectivePermissions"];
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.breakRoleInheritance = _funcs__WEBPACK_IMPORTED_MODULE_4__["breakRoleInheritance"];
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.resetRoleInheritance = _funcs__WEBPACK_IMPORTED_MODULE_4__["resetRoleInheritance"];
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.userHasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["userHasPermissions"];
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.currentUserHasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["currentUserHasPermissions"];
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.hasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["hasPermissions"];
//# sourceMappingURL=item.js.map

/***/ }),

/***/ "Ku5p":
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/sputilities/index.js ***!
  \***************************************************/
/*! exports provided: Utilities */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "AfLD");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Utilities", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Utilities"]; });




Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"].prototype, "utility", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_1__["Utilities"])(this._baseUrl, "");
    },
});
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "KvhF":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/comments/index.js ***!
  \************************************************/
/*! exports provided: Comment, Comments, Replies */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _item__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./item */ "NxAp");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "wMck");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Comment", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Comment"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Comments", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Comments"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Replies", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Replies"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "LVfT":
/*!**************************************************************!*\
  !*** ./node_modules/@pnp/sp/node_modules/tslib/tslib.es6.js ***!
  \**************************************************************/
/*! exports provided: __extends, __assign, __rest, __decorate, __param, __metadata, __awaiter, __generator, __exportStar, __values, __read, __spread, __spreadArrays, __await, __asyncGenerator, __asyncDelegator, __asyncValues, __makeTemplateObject, __importStar, __importDefault */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__extends", function() { return __extends; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__assign", function() { return __assign; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__rest", function() { return __rest; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__decorate", function() { return __decorate; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__param", function() { return __param; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__metadata", function() { return __metadata; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__awaiter", function() { return __awaiter; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__generator", function() { return __generator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__exportStar", function() { return __exportStar; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__values", function() { return __values; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__read", function() { return __read; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__spread", function() { return __spread; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__spreadArrays", function() { return __spreadArrays; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__await", function() { return __await; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncGenerator", function() { return __asyncGenerator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncDelegator", function() { return __asyncDelegator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncValues", function() { return __asyncValues; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__makeTemplateObject", function() { return __makeTemplateObject; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__importStar", function() { return __importStar; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__importDefault", function() { return __importDefault; });
/*! *****************************************************************************
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use
this file except in compliance with the License. You may obtain a copy of the
License at http://www.apache.org/licenses/LICENSE-2.0

THIS CODE IS PROVIDED ON AN *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, EITHER EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION ANY IMPLIED
WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE,
MERCHANTABLITY OR NON-INFRINGEMENT.

See the Apache Version 2.0 License for specific language governing permissions
and limitations under the License.
***************************************************************************** */
/* global Reflect, Promise */

var extendStatics = function(d, b) {
    extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return extendStatics(d, b);
};

function __extends(d, b) {
    extendStatics(d, b);
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
}

var __assign = function() {
    __assign = Object.assign || function __assign(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
        }
        return t;
    }
    return __assign.apply(this, arguments);
}

function __rest(s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
}

function __decorate(decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
}

function __param(paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
}

function __metadata(metadataKey, metadataValue) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(metadataKey, metadataValue);
}

function __awaiter(thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
}

function __generator(thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
}

function __exportStar(m, exports) {
    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
}

function __values(o) {
    var m = typeof Symbol === "function" && o[Symbol.iterator], i = 0;
    if (m) return m.call(o);
    return {
        next: function () {
            if (o && i >= o.length) o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
}

function __read(o, n) {
    var m = typeof Symbol === "function" && o[Symbol.iterator];
    if (!m) return o;
    var i = m.call(o), r, ar = [], e;
    try {
        while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
    }
    catch (error) { e = { error: error }; }
    finally {
        try {
            if (r && !r.done && (m = i["return"])) m.call(i);
        }
        finally { if (e) throw e.error; }
    }
    return ar;
}

function __spread() {
    for (var ar = [], i = 0; i < arguments.length; i++)
        ar = ar.concat(__read(arguments[i]));
    return ar;
}

function __spreadArrays() {
    for (var s = 0, i = 0, il = arguments.length; i < il; i++) s += arguments[i].length;
    for (var r = Array(s), k = 0, i = 0; i < il; i++)
        for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
            r[k] = a[j];
    return r;
};

function __await(v) {
    return this instanceof __await ? (this.v = v, this) : new __await(v);
}

function __asyncGenerator(thisArg, _arguments, generator) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var g = generator.apply(thisArg, _arguments || []), i, q = [];
    return i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i;
    function verb(n) { if (g[n]) i[n] = function (v) { return new Promise(function (a, b) { q.push([n, v, a, b]) > 1 || resume(n, v); }); }; }
    function resume(n, v) { try { step(g[n](v)); } catch (e) { settle(q[0][3], e); } }
    function step(r) { r.value instanceof __await ? Promise.resolve(r.value.v).then(fulfill, reject) : settle(q[0][2], r); }
    function fulfill(value) { resume("next", value); }
    function reject(value) { resume("throw", value); }
    function settle(f, v) { if (f(v), q.shift(), q.length) resume(q[0][0], q[0][1]); }
}

function __asyncDelegator(o) {
    var i, p;
    return i = {}, verb("next"), verb("throw", function (e) { throw e; }), verb("return"), i[Symbol.iterator] = function () { return this; }, i;
    function verb(n, f) { i[n] = o[n] ? function (v) { return (p = !p) ? { value: __await(o[n](v)), done: n === "return" } : f ? f(v) : v; } : f; }
}

function __asyncValues(o) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var m = o[Symbol.asyncIterator], i;
    return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
    function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
    function settle(resolve, reject, d, v) { Promise.resolve(v).then(function(v) { resolve({ value: v, done: d }); }, reject); }
}

function __makeTemplateObject(cooked, raw) {
    if (Object.defineProperty) { Object.defineProperty(cooked, "raw", { value: raw }); } else { cooked.raw = raw; }
    return cooked;
};

function __importStar(mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result.default = mod;
    return result;
}

function __importDefault(mod) {
    return (mod && mod.__esModule) ? mod : { default: mod };
}


/***/ }),

/***/ "LZ64":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/item.js ***!
  \**********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "4tXH");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"], "folder", _types__WEBPACK_IMPORTED_MODULE_2__["Folder"], "folder");
//# sourceMappingURL=item.js.map

/***/ }),

/***/ "M9RU":
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/web.js ***!
  \***************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "S3gC");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "contentTypes", _types__WEBPACK_IMPORTED_MODULE_2__["ContentTypes"]);
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "MkPD":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/search/query.js ***!
  \**********************************************/
/*! exports provided: SearchQueryBuilder, _Search, Search, SearchResults */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SearchQueryBuilder", function() { return SearchQueryBuilder; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Search", function() { return _Search; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Search", function() { return Search; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SearchResults", function() { return SearchResults; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../telemetry */ "h3gW");








var funcs = new Map([
    ["text", "Querytext"],
    ["template", "QueryTemplate"],
    ["sourceId", "SourceId"],
    ["trimDuplicatesIncludeId", ""],
    ["startRow", ""],
    ["rowLimit", ""],
    ["rankingModelId", ""],
    ["rowsPerPage", ""],
    ["selectProperties", ""],
    ["culture", ""],
    ["timeZoneId", ""],
    ["refinementFilters", ""],
    ["refiners", ""],
    ["hiddenConstraints", ""],
    ["sortList", ""],
    ["timeout", ""],
    ["hithighlightedProperties", ""],
    ["clientType", ""],
    ["personalizationData", ""],
    ["resultsURL", ""],
    ["queryTag", ""],
    ["properties", ""],
    ["queryTemplatePropertiesUrl", ""],
    ["reorderingRules", ""],
    ["hitHighlightedMultivaluePropertyLimit", ""],
    ["collapseSpecification", ""],
    ["uiLanguage", ""],
    ["desiredSnippetLength", ""],
    ["maxSnippetLength", ""],
    ["summaryLength", ""],
]);
var props = new Map([]);
function toPropCase(str) {
    return str.replace(/^(.)/, function ($1) { return $1.toUpperCase(); });
}
/**
 * Creates a new instance of the SearchQueryBuilder
 *
 * @param queryText Initial query text
 * @param _query Any initial query configuration
 */
function SearchQueryBuilder(queryText, _query) {
    if (queryText === void 0) { queryText = ""; }
    if (_query === void 0) { _query = {}; }
    return new Proxy({
        query: Object.assign({
            Querytext: queryText,
        }, _query),
    }, {
        get: function (self, propertyKey, proxy) {
            var pk = propertyKey.toString();
            if (pk === "toSearchQuery") {
                return function () { return self.query; };
            }
            if (funcs.has(pk)) {
                return function () {
                    var value = [];
                    for (var _i = 0; _i < arguments.length; _i++) {
                        value[_i] = arguments[_i];
                    }
                    var mappedPk = funcs.get(pk);
                    self.query[mappedPk.length > 0 ? mappedPk : toPropCase(pk)] = value.length > 1 ? value : value[0];
                    return proxy;
                };
            }
            var propKey = props.has(pk) ? props.get(pk) : toPropCase(pk);
            self.query[propKey] = true;
            return proxy;
        },
    });
}
var queryRegex = /_api\/search\/postquery$/i;
/**
 * Describes the search API
 *
 */
var _Search = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Search, _super);
    function _Search() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * @returns Promise
     */
    _Search.prototype.execute = function (queryInit) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var query, postBody, cacheKey, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        query = this.parseQuery(queryInit);
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_4__["body"])({
                            request: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_3__["metadata"])("Microsoft.Office.Server.Search.REST.SearchRequest"), Object.assign({}, query, {
                                HitHighlightedProperties: this.fixArrProp(query.HitHighlightedProperties),
                                Properties: this.fixArrProp(query.Properties),
                                RefinementFilters: this.fixArrProp(query.RefinementFilters),
                                ReorderingRules: this.fixArrProp(query.ReorderingRules),
                                SelectProperties: this.fixArrProp(query.SelectProperties),
                                SortList: this.fixArrProp(query.SortList),
                            })),
                        });
                        // if we are using caching with this search request, then we need to handle some work upfront to enable that
                        if (this.data.useCaching) {
                            // force use of the cache for this request if .usingCaching was called
                            this._forceCaching = true;
                            cacheKey = "PnPjs.SearchWithCaching(" + Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["getHashCode"])(postBody.body) + ")";
                            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["objectDefinedNotNull"])(this.data.cachingOptions)) {
                                // if our key ends in the postquery url we overwrite it
                                if (queryRegex.test(this.data.cachingOptions.key)) {
                                    this.data.cachingOptions.key = cacheKey;
                                }
                            }
                            else {
                                this.data.cachingOptions = new _pnp_odata__WEBPACK_IMPORTED_MODULE_4__["CachingOptions"](cacheKey);
                            }
                        }
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this, postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, new SearchResults(data, this.toUrl(), query)];
                }
            });
        });
    };
    /**
     * Fix array property
     *
     * @param prop property to fix for container struct
     */
    _Search.prototype.fixArrProp = function (prop) {
        if (typeof prop === "undefined") {
            return ({ results: [] });
        }
        return { results: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["isArray"])(prop) ? prop : [prop] };
    };
    /**
     * Translates one of the query initializers into a SearchQuery instance
     *
     * @param query
     */
    _Search.prototype.parseQuery = function (query) {
        var finalQuery;
        if (typeof query === "string") {
            finalQuery = { Querytext: query };
        }
        else if (query.toSearchQuery) {
            finalQuery = query.toSearchQuery();
        }
        else {
            finalQuery = query;
        }
        return finalQuery;
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_7__["tag"])("se.execute")
    ], _Search.prototype, "execute", null);
    _Search = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("_api/search/postquery")
    ], _Search);
    return _Search;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Search = function (baseUrl, options) {
    if (options === void 0) { options = {}; }
    return function (queryInit) {
        return (new _Search(baseUrl)).configure(options).execute(queryInit);
    };
};
var SearchResults = /** @class */ (function () {
    function SearchResults(rawResponse, _url, _query, _raw, _primary) {
        if (_raw === void 0) { _raw = null; }
        if (_primary === void 0) { _primary = null; }
        this._url = _url;
        this._query = _query;
        this._raw = _raw;
        this._primary = _primary;
        this._url = this._url.replace(queryRegex, "");
        this._raw = rawResponse.postquery ? rawResponse.postquery : rawResponse;
    }
    Object.defineProperty(SearchResults.prototype, "ElapsedTime", {
        get: function () {
            var _a, _b;
            return ((_b = (_a = this) === null || _a === void 0 ? void 0 : _a.RawSearchResults) === null || _b === void 0 ? void 0 : _b.ElapsedTime) || 0;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SearchResults.prototype, "RowCount", {
        get: function () {
            var _a, _b, _c, _d;
            return ((_d = (_c = (_b = (_a = this) === null || _a === void 0 ? void 0 : _a.RawSearchResults) === null || _b === void 0 ? void 0 : _b.PrimaryQueryResult) === null || _c === void 0 ? void 0 : _c.RelevantResults) === null || _d === void 0 ? void 0 : _d.RowCount) || 0;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SearchResults.prototype, "TotalRows", {
        get: function () {
            var _a, _b, _c, _d;
            return ((_d = (_c = (_b = (_a = this) === null || _a === void 0 ? void 0 : _a.RawSearchResults) === null || _b === void 0 ? void 0 : _b.PrimaryQueryResult) === null || _c === void 0 ? void 0 : _c.RelevantResults) === null || _d === void 0 ? void 0 : _d.TotalRows) || 0;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SearchResults.prototype, "TotalRowsIncludingDuplicates", {
        get: function () {
            var _a, _b, _c, _d;
            return ((_d = (_c = (_b = (_a = this) === null || _a === void 0 ? void 0 : _a.RawSearchResults) === null || _b === void 0 ? void 0 : _b.PrimaryQueryResult) === null || _c === void 0 ? void 0 : _c.RelevantResults) === null || _d === void 0 ? void 0 : _d.TotalRowsIncludingDuplicates) || 0;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SearchResults.prototype, "RawSearchResults", {
        get: function () {
            return this._raw;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SearchResults.prototype, "PrimarySearchResults", {
        get: function () {
            var _a, _b, _c, _d;
            if (this._primary === null) {
                this._primary = this.formatSearchResults(((_d = (_c = (_b = (_a = this._raw) === null || _a === void 0 ? void 0 : _a.PrimaryQueryResult) === null || _b === void 0 ? void 0 : _b.RelevantResults) === null || _c === void 0 ? void 0 : _c.Table) === null || _d === void 0 ? void 0 : _d.Rows) || null);
            }
            return this._primary;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Gets a page of results
     *
     * @param pageNumber Index of the page to return. Used to determine StartRow
     * @param pageSize Optional, items per page (default = 10)
     */
    SearchResults.prototype.getPage = function (pageNumber, pageSize) {
        // if we got all the available rows we don't have another page
        if (this.TotalRows < this.RowCount) {
            return Promise.resolve(null);
        }
        // if pageSize is supplied, then we use that regardless of any previous values
        // otherwise get the previous RowLimit or default to 10
        var rows = pageSize !== undefined ? pageSize : Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(this._query, "RowLimit") ? this._query.RowLimit : 10;
        var query = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(this._query, {
            RowLimit: rows,
            StartRow: rows * (pageNumber - 1),
        });
        // we have reached the end
        if (query.StartRow > this.TotalRows) {
            return Promise.resolve(null);
        }
        return Search(this._url)(query);
    };
    /**
     * Formats a search results array
     *
     * @param rawResults The array to process
     */
    SearchResults.prototype.formatSearchResults = function (rawResults) {
        var e_1, _a;
        var results = new Array();
        if (typeof (rawResults) === "undefined" || rawResults == null) {
            return [];
        }
        var tempResults = rawResults.results ? rawResults.results : rawResults;
        try {
            for (var tempResults_1 = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__values"])(tempResults), tempResults_1_1 = tempResults_1.next(); !tempResults_1_1.done; tempResults_1_1 = tempResults_1.next()) {
                var tempResult = tempResults_1_1.value;
                var cells = tempResult.Cells.results ? tempResult.Cells.results : tempResult.Cells;
                results.push(cells.reduce(function (res, cell) {
                    Reflect.defineProperty(res, cell.Key, {
                        configurable: false,
                        enumerable: true,
                        value: cell.Value,
                        writable: false,
                    });
                    return res;
                }, {}));
            }
        }
        catch (e_1_1) { e_1 = { error: e_1_1 }; }
        finally {
            try {
                if (tempResults_1_1 && !tempResults_1_1.done && (_a = tempResults_1.return)) _a.call(tempResults_1);
            }
            finally { if (e_1) throw e_1.error; }
        }
        return results;
    };
    return SearchResults;
}());

//# sourceMappingURL=query.js.map

/***/ }),

/***/ "N3o1":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/subscriptions/types.js ***!
  \*****************************************************/
/*! exports provided: _Subscriptions, Subscriptions, _Subscription, Subscription */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Subscriptions", function() { return _Subscriptions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Subscriptions", function() { return Subscriptions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Subscription", function() { return _Subscription; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Subscription", function() { return Subscription; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../telemetry */ "h3gW");






var _Subscriptions = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Subscriptions, _super);
    function _Subscriptions() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
    * Returns all the webhook subscriptions or the specified webhook subscription
    *
    * @param subscriptionId The id of a specific webhook subscription to retrieve, omit to retrieve all the webhook subscriptions
    */
    _Subscriptions.prototype.getById = function (subscriptionId) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"].configure(Subscription(this).concat("('" + subscriptionId + "')"), "subs.getById");
    };
    /**
     * Creates a new webhook subscription
     *
     * @param notificationUrl The url to receive the notifications
     * @param expirationDate The date and time to expire the subscription in the form YYYY-MM-ddTHH:mm:ss+00:00 (maximum of 6 months)
     * @param clientState A client specific string (optional)
     */
    _Subscriptions.prototype.add = function (notificationUrl, expirationDate, clientState) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = {
                            "expirationDateTime": expirationDate,
                            "notificationUrl": notificationUrl,
                            "resource": this.toUrl(),
                        };
                        if (clientState) {
                            postBody.clientState = clientState;
                        }
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(this, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])(postBody, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "Content-Type": "application/json" })))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, { data: data, subscription: this.getById(data.id) }];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("subs.add")
    ], _Subscriptions.prototype, "add", null);
    _Subscriptions = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("subscriptions")
    ], _Subscriptions);
    return _Subscriptions;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var Subscriptions = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Subscriptions);
var _Subscription = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Subscription, _super);
    function _Subscription() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Renews this webhook subscription
     *
     * @param expirationDate The date and time to expire the subscription in the form YYYY-MM-ddTHH:mm:ss+00:00 (maximum of 6 months, optional)
     * @param notificationUrl The url to receive the notifications (optional)
     * @param clientState A client specific string (optional)
     */
    _Subscription.prototype.update = function (expirationDate, notificationUrl, clientState) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = {};
                        if (expirationDate) {
                            postBody.expirationDateTime = expirationDate;
                        }
                        if (notificationUrl) {
                            postBody.notificationUrl = notificationUrl;
                        }
                        if (clientState) {
                            postBody.clientState = clientState;
                        }
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPatch"])(this, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])(postBody, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "Content-Type": "application/json" })))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, { data: data, subscription: this }];
                }
            });
        });
    };
    /**
     * Removes this webhook subscription
     *
     */
    _Subscription.prototype.delete = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spDelete"])(this);
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("sub.update")
    ], _Subscription.prototype, "update", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("sub.delete")
    ], _Subscription.prototype, "delete", null);
    return _Subscription;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Subscription = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Subscription);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "NCg5":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/site-scripts/web.js ***!
  \**************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "b9JD");


_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getSiteScript = function (extractInfo) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["SiteScripts"])(this.toUrl(), "").getSiteScriptFromWeb(undefined, extractInfo);
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "NTTg":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/items/list.js ***!
  \********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "3DT9");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "items", _types__WEBPACK_IMPORTED_MODULE_2__["Items"]);
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "NxAp":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/comments/item.js ***!
  \***********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "wMck");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../operations */ "UK2s");




Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"], "comments", _types__WEBPACK_IMPORTED_MODULE_2__["Comments"]);
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.getLikedBy = function () {
    return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(_items_types__WEBPACK_IMPORTED_MODULE_1__["Item"], "likedBy"));
};
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.like = function () {
    return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(_items_types__WEBPACK_IMPORTED_MODULE_1__["Item"], "like"));
};
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.unlike = function () {
    return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(_items_types__WEBPACK_IMPORTED_MODULE_1__["Item"], "unlike"));
};
_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"].prototype.getLikedByInformation = function () {
    return this.clone(_items_types__WEBPACK_IMPORTED_MODULE_1__["Item"], "likedByInformation").expand("likedby")();
};
//# sourceMappingURL=item.js.map

/***/ }),

/***/ "P6NM":
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/column-defaults/folder.js ***!
  \********************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _lists_web__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../lists/web */ "Bwa7");
/* harmony import */ var _folders_types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../folders/types */ "4tXH");






_folders_types__WEBPACK_IMPORTED_MODULE_5__["_Folder"].prototype.getDefaultColumnValues = function () {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var folderProps, serRelPath, web, docLib;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_folders_types__WEBPACK_IMPORTED_MODULE_5__["Folder"])(this, "Properties").select("vti_x005f_listname")()];
                case 1:
                    folderProps = _a.sent();
                    return [4 /*yield*/, this.select("ServerRelativePath")()];
                case 2:
                    serRelPath = (_a.sent()).ServerRelativePath;
                    web = Object(_webs_types__WEBPACK_IMPORTED_MODULE_3__["Web"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_2__["extractWebUrl"])(Object(_odata__WEBPACK_IMPORTED_MODULE_1__["odataUrlFrom"])(folderProps)));
                    docLib = web.lists.getById(folderProps.vti_x005f_listname);
                    return [4 /*yield*/, docLib.getDefaultColumnValues()];
                case 3: 
                // and we return the defaults associated with this folder's server relative path only
                // if you want all the defaults use list.getDefaultColumnValues()
                return [2 /*return*/, (_a.sent()).filter(function (v) { return v.path.toLowerCase() === serRelPath.DecodedUrl.toLowerCase(); })];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_5__["_Folder"].prototype.setDefaultColumnValues = function (fieldDefaults, merge) {
    if (merge === void 0) { merge = true; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var folderProps, web, docLib, folderPath, existingDefaults, filteredExistingDefaults;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_folders_types__WEBPACK_IMPORTED_MODULE_5__["Folder"])(this, "Properties").select("vti_x005f_listname")()];
                case 1:
                    folderProps = _a.sent();
                    web = Object(_webs_types__WEBPACK_IMPORTED_MODULE_3__["Web"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_2__["extractWebUrl"])(Object(_odata__WEBPACK_IMPORTED_MODULE_1__["odataUrlFrom"])(folderProps)));
                    docLib = web.lists.getById(folderProps.vti_x005f_listname);
                    return [4 /*yield*/, this.select("ServerRelativePath")()];
                case 2:
                    folderPath = (_a.sent()).ServerRelativePath.DecodedUrl;
                    return [4 /*yield*/, docLib.getDefaultColumnValues()];
                case 3:
                    existingDefaults = _a.sent();
                    filteredExistingDefaults = merge ? existingDefaults : existingDefaults.filter(function (f) { return f.path !== folderPath; });
                    // we update / add any new defaults from those passed to this method
                    fieldDefaults.forEach(function (d) {
                        var existing = filteredExistingDefaults.find(function (ed) { return ed.name === d.name && ed.path === folderPath; });
                        if (existing) {
                            existing.value = d.value;
                        }
                        else {
                            filteredExistingDefaults.push({
                                name: d.name,
                                path: folderPath,
                                value: d.value,
                            });
                        }
                    });
                    // after this operation filteredExistingDefaults should contain all the value we want to write to the file
                    return [4 /*yield*/, docLib.setDefaultColumnValues(filteredExistingDefaults)];
                case 4:
                    // after this operation filteredExistingDefaults should contain all the value we want to write to the file
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_5__["_Folder"].prototype.clearDefaultColumnValues = function () {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, this.setDefaultColumnValues([], false)];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
};
//# sourceMappingURL=folder.js.map

/***/ }),

/***/ "PmFa":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/related-items/types.js ***!
  \*****************************************************/
/*! exports provided: _RelatedItemManager, RelatedItemManager */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_RelatedItemManager", function() { return _RelatedItemManager; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RelatedItemManager", function() { return RelatedItemManager; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../telemetry */ "h3gW");







var _RelatedItemManager = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_RelatedItemManager, _super);
    function _RelatedItemManager() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    _RelatedItemManager.prototype.getRelatedItems = function (sourceListName, sourceItemId) {
        var query = this.clone(RelatedItemManager, null);
        query.concat(".GetRelatedItems");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(query, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
        }));
    };
    _RelatedItemManager.prototype.getPageOneRelatedItems = function (sourceListName, sourceItemId) {
        var query = this.clone(RelatedItemManager, null);
        query.concat(".GetPageOneRelatedItems");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(query, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
        }));
    };
    _RelatedItemManager.prototype.addSingleLink = function (sourceListName, sourceItemId, sourceWebUrl, targetListName, targetItemID, targetWebUrl, tryAddReverseLink) {
        if (tryAddReverseLink === void 0) { tryAddReverseLink = false; }
        var query = this.clone(RelatedItemManager, null);
        query.concat(".AddSingleLink");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(query, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
            SourceWebUrl: sourceWebUrl,
            TargetItemID: targetItemID,
            TargetListName: targetListName,
            TargetWebUrl: targetWebUrl,
            TryAddReverseLink: tryAddReverseLink,
        }));
    };
    _RelatedItemManager.prototype.addSingleLinkToUrl = function (sourceListName, sourceItemId, targetItemUrl, tryAddReverseLink) {
        if (tryAddReverseLink === void 0) { tryAddReverseLink = false; }
        var query = this.clone(RelatedItemManager, null);
        query.concat(".AddSingleLinkToUrl");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(query, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
            TargetItemUrl: targetItemUrl,
            TryAddReverseLink: tryAddReverseLink,
        }));
    };
    _RelatedItemManager.prototype.addSingleLinkFromUrl = function (sourceItemUrl, targetListName, targetItemId, tryAddReverseLink) {
        if (tryAddReverseLink === void 0) { tryAddReverseLink = false; }
        var query = this.clone(RelatedItemManager, null);
        query.concat(".AddSingleLinkFromUrl");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(query, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
            SourceItemUrl: sourceItemUrl,
            TargetItemID: targetItemId,
            TargetListName: targetListName,
            TryAddReverseLink: tryAddReverseLink,
        }));
    };
    _RelatedItemManager.prototype.deleteSingleLink = function (sourceListName, sourceItemId, sourceWebUrl, targetListName, targetItemId, targetWebUrl, tryDeleteReverseLink) {
        if (tryDeleteReverseLink === void 0) { tryDeleteReverseLink = false; }
        var query = this.clone(RelatedItemManager, null);
        query.concat(".DeleteSingleLink");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(query, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
            SourceItemID: sourceItemId,
            SourceListName: sourceListName,
            SourceWebUrl: sourceWebUrl,
            TargetItemID: targetItemId,
            TargetListName: targetListName,
            TargetWebUrl: targetWebUrl,
            TryDeleteReverseLink: tryDeleteReverseLink,
        }));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("rim.getRelatedItems")
    ], _RelatedItemManager.prototype, "getRelatedItems", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("rim.getPageOneRelatedItems")
    ], _RelatedItemManager.prototype, "getPageOneRelatedItems", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("rim.addSingleLink")
    ], _RelatedItemManager.prototype, "addSingleLink", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("rim.ToUrl")
    ], _RelatedItemManager.prototype, "addSingleLinkToUrl", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("rim.FromUrl")
    ], _RelatedItemManager.prototype, "addSingleLinkFromUrl", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("rim.deleteSingleLink")
    ], _RelatedItemManager.prototype, "deleteSingleLink", null);
    _RelatedItemManager = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("_api/SP.RelatedItemManager")
    ], _RelatedItemManager);
    return _RelatedItemManager;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryable"]));

var RelatedItemManager = function (url) { return new _RelatedItemManager(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_2__["extractWebUrl"])(typeof url === "string" ? url : url.toUrl())); };
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "S3gC":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/types.js ***!
  \*****************************************************/
/*! exports provided: _ContentTypes, ContentTypes, _ContentType, ContentType, _FieldLinks, FieldLinks, _FieldLink, FieldLink */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_ContentTypes", function() { return _ContentTypes; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ContentTypes", function() { return ContentTypes; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_ContentType", function() { return _ContentType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ContentType", function() { return ContentType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_FieldLinks", function() { return _FieldLinks; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "FieldLinks", function() { return FieldLinks; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_FieldLink", function() { return _FieldLink; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "FieldLink", function() { return FieldLink; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../telemetry */ "h3gW");







var _ContentTypes = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_ContentTypes, _super);
    function _ContentTypes() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Adds an existing contenttype to a content type collection
     *
     * @param contentTypeId in the following format, for example: 0x010102
     */
    _ContentTypes.prototype.addAvailableContentType = function (contentTypeId) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(ContentTypes, "addAvailableContentType"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({ "contentTypeId": contentTypeId }))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                contentType: this.getById(data.id),
                                data: data,
                            }];
                }
            });
        });
    };
    /**
     * Gets a ContentType by content type id
     * @param id The id of the content type to get, in the following format, for example: 0x010102
     */
    _ContentTypes.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(ContentType(this).concat("('" + id + "')"), "cts.getById");
    };
    /**
     * Adds a new content type to the collection
     *
     * @param id The desired content type id for the new content type (also determines the parent content type)
     * @param name The name of the content type
     * @param description The description of the content type
     * @param group The group in which to add the content type
     * @param additionalSettings Any additional settings to provide when creating the content type
     *
     */
    _ContentTypes.prototype.add = function (id, name, description, group, additionalSettings) {
        if (description === void 0) { description = ""; }
        if (group === void 0) { group = "Custom Content Types"; }
        if (additionalSettings === void 0) { additionalSettings = {}; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(Object.assign(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_4__["metadata"])("SP.ContentType"), {
                            "Description": description,
                            "Group": group,
                            "Id": { "StringValue": id },
                            "Name": name,
                        }, additionalSettings));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this, postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, { contentType: this.getById(data.id), data: data }];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("cts.addAvailableContentType")
    ], _ContentTypes.prototype, "addAvailableContentType", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("cts.add")
    ], _ContentTypes.prototype, "add", null);
    _ContentTypes = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("contenttypes")
    ], _ContentTypes);
    return _ContentTypes;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var ContentTypes = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_ContentTypes);
var _ContentType = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_ContentType, _super);
    function _ContentType() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["deleteable"])("ct");
        return _this;
    }
    Object.defineProperty(_ContentType.prototype, "fieldLinks", {
        /**
         * Gets the column (also known as field) references in the content type.
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(FieldLinks(this), "ct.fieldLinks");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ContentType.prototype, "fields", {
        /**
         * Gets a value that specifies the collection of fields for the content type.
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableCollection"])(this, "fields"), "ct.fields");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ContentType.prototype, "parent", {
        /**
         * Gets the parent content type of the content type.
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(ContentType(this, "parent"), "ct.parent");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ContentType.prototype, "workflowAssociations", {
        /**
         * Gets a value that specifies the collection of workflow associations for the content type.
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableCollection"])(this, "workflowAssociations"), "ct.workflowAssociations");
        },
        enumerable: true,
        configurable: true
    });
    return _ContentType;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableInstance"]));

var ContentType = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_ContentType);
var _FieldLinks = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_FieldLinks, _super);
    function _FieldLinks() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
    *  Gets a FieldLink by GUID id
    *
    * @param id The GUID id of the field link
    */
    _FieldLinks.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(FieldLink(this).concat("(guid'" + id + "')"), "fls.getById");
    };
    _FieldLinks = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("fieldlinks")
    ], _FieldLinks);
    return _FieldLinks;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var FieldLinks = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_FieldLinks);
var _FieldLink = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_FieldLink, _super);
    function _FieldLink() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return _FieldLink;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableInstance"]));

var FieldLink = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_FieldLink);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "SPHc":
/*!*******************************************!*\
  !*** ./node_modules/@pnp/sp/files/web.js ***!
  \*******************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "qfcF");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");



_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getFileByServerRelativeUrl = function (fileRelativeUrl) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["File"])(this, "getFileByServerRelativeUrl('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_2__["escapeQueryStrValue"])(fileRelativeUrl) + "')");
};
_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getFileByServerRelativePath = function (fileRelativeUrl) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["File"])(this, "getFileByServerRelativePath(decodedUrl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_2__["escapeQueryStrValue"])(fileRelativeUrl) + "')");
};
_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getFileById = function (uniqueId) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["File"])(this, "getFileById('" + uniqueId + "')");
};
_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getFileByUrl = function (fileUrl) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["File"])(this, "getFileByUrl('!@p1::" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_2__["escapeQueryStrValue"])(fileUrl) + "')");
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "SQg3":
/*!******************************************!*\
  !*** ./node_modules/@pnp/odata/batch.js ***!
  \******************************************/
/*! exports provided: Batch */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Batch", function() { return Batch; });
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/common */ "zqEK");

var Batch = /** @class */ (function () {
    function Batch(_batchId) {
        if (_batchId === void 0) { _batchId = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["getGUID"])(); }
        this._batchId = _batchId;
        this._reqs = [];
        this._deps = [];
        this._rDeps = [];
        this._index = -1;
    }
    Object.defineProperty(Batch.prototype, "batchId", {
        get: function () {
            return this._batchId;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Batch.prototype, "requests", {
        /**
         * The requests contained in this batch
         */
        get: function () {
            // we sort these each time this is accessed
            return this._reqs.sort(function (info1, info2) { return info1.index - info2.index; });
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Not meant for use directly
     *
     * @param batchee The IQueryable for this batch to track in order
     */
    Batch.prototype.track = function (batchee) {
        batchee.data.batch = this;
        // we need to track the order requests are added to the batch to ensure we always
        // operate on them in order
        if (typeof batchee.data.batchIndex === "undefined" || batchee.data.batchIndex < 0) {
            batchee.data.batchIndex = ++this._index;
        }
    };
    /**
     * Adds the given request context to the batch for execution
     *
     * @param context Details of the request to batch
     */
    Batch.prototype.add = function (context) {
        var info = {
            id: context.requestId,
            index: context.batchIndex,
            method: context.method.toUpperCase(),
            options: context.options,
            parser: context.parser,
            reject: null,
            resolve: null,
            url: context.url,
        };
        // we create a new promise that will be resolved within the batch
        var p = new Promise(function (resolve, reject) {
            info.resolve = resolve;
            info.reject = reject;
        });
        this._reqs.push(info);
        return p;
    };
    /**
     * Adds a dependency insuring that some set of actions will occur before a batch is processed.
     * MUST be cleared using the returned resolve delegate to allow batches to run
     */
    Batch.prototype.addDependency = function () {
        var resolver = function () { return void (0); };
        this._deps.push(new Promise(function (resolve) {
            resolver = resolve;
        }));
        return resolver;
    };
    /**
     * The batch's execute method will not resolve util any promises added here resolve
     *
     * @param p The dependent promise
     */
    Batch.prototype.addResolveBatchDependency = function (p) {
        this._rDeps.push(p);
    };
    /**
     * Execute the current batch and resolve the associated promises
     *
     * @returns A promise which will be resolved once all of the batch's child promises have resolved
     */
    Batch.prototype.execute = function () {
        var _this = this;
        // we need to check the dependencies twice due to how different engines handle things.
        // We can get a second set of promises added during the first set resolving
        return Promise.all(this._deps)
            .then(function () { return Promise.all(_this._deps); })
            .then(function () { return _this.executeImpl(); })
            .then(function () { return Promise.all(_this._rDeps); })
            .then(function () { return void (0); });
    };
    return Batch;
}());

//# sourceMappingURL=batch.js.map

/***/ }),

/***/ "Sdn1":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/hubsites/index.js ***!
  \************************************************/
/*! exports provided: HubSite, HubSites */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "WqOe");
/* harmony import */ var _site__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./site */ "EyfK");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./web */ "37pO");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "HubSite", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["HubSite"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "HubSites", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["HubSites"]; });






Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"].prototype, "hubSites", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_1__["HubSites"])(this._baseUrl).configure(this._options);
    },
});
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "T2j4":
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/funcs.js ***!
  \*********************************************************/
/*! exports provided: getValueForUICultureBinder */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getValueForUICultureBinder", function() { return getValueForUICultureBinder; });
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");



function getValueForUICultureBinder(propName) {
    return function (cultureName) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_1__["spPost"])(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_0__["SharePointQueryable"], propName + "/getValueForUICulture"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({ cultureName: cultureName }));
    };
}
//# sourceMappingURL=funcs.js.map

/***/ }),

/***/ "TUji":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/related-items/index.js ***!
  \*****************************************************/
/*! exports provided: RelatedItemManager */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./types */ "PmFa");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RelatedItemManager", function() { return _types__WEBPACK_IMPORTED_MODULE_0__["RelatedItemManager"]; });


//# sourceMappingURL=index.js.map

/***/ }),

/***/ "TZA7":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/odata/pipeline.js ***!
  \*********************************************/
/*! exports provided: setResult, pipe, requestPipelineMethod, PipelineMethods, getDefaultPipeline */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setResult", function() { return setResult; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "pipe", function() { return pipe; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "requestPipelineMethod", function() { return requestPipelineMethod; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PipelineMethods", function() { return PipelineMethods; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getDefaultPipeline", function() { return getDefaultPipeline; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "t3D1");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_logging__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/logging */ "B3L+");
/* harmony import */ var _caching__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./caching */ "Zl34");




/**
 * Resolves the context's result value
 *
 * @param context The current context
 */
function returnResult(context) {
    _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].log({
        data: _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].activeLogLevel === 0 /* Verbose */ ? context.result : {},
        level: 1 /* Info */,
        message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Returning result from pipeline. Set logging to verbose to see data.",
    });
    return Promise.resolve(context.result);
}
/**
 * Sets the result on the context
 */
function setResult(context, value) {
    return new Promise(function (resolve) {
        context.result = value;
        context.hasResult = true;
        resolve(context);
    });
}
/**
 * Invokes the next method in the provided context's pipeline
 *
 * @param c The current request context
 */
function next(c) {
    return c.pipes.length > 0 ? c.pipes.shift()(c) : Promise.resolve(c);
}
/**
 * Executes the current request context's pipeline
 *
 * @param context Current context
 */
function pipe(context) {
    if (context.pipes.length < 1) {
        _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Request pipeline contains no methods!", 3 /* Error */);
        throw Error("Request pipeline contains no methods!");
    }
    var promise = next(context).then(function (ctx) { return returnResult(ctx); }).catch(function (e) {
        _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].error(e);
        throw e;
    });
    if (context.isBatched) {
        // this will block the batch's execute method from returning until the child requets have been resolved
        context.batch.addResolveBatchDependency(promise);
    }
    return promise;
}
/**
 * decorator factory applied to methods in the pipeline to control behavior
 */
function requestPipelineMethod(alwaysRun) {
    if (alwaysRun === void 0) { alwaysRun = false; }
    return function (target, propertyKey, descriptor) {
        var method = descriptor.value;
        descriptor.value = function () {
            var args = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                args[_i] = arguments[_i];
            }
            // if we have a result already in the pipeline, pass it along and don't call the tagged method
            if (!alwaysRun && args.length > 0 && Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(args[0], "hasResult") && args[0].hasResult) {
                _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].write("[" + args[0].requestId + "] (" + (new Date()).getTime() + ") Skipping request pipeline method " + propertyKey + ", existing result in pipeline.", 0 /* Verbose */);
                return Promise.resolve(args[0]);
            }
            // apply the tagged method
            _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].write("[" + args[0].requestId + "] (" + (new Date()).getTime() + ") Calling request pipeline method " + propertyKey + ".", 0 /* Verbose */);
            // then chain the next method in the context's pipeline - allows for dynamic pipeline
            return method.apply(target, args).then(function (ctx) { return next(ctx); });
        };
    };
}
/**
 * Contains the methods used within the request pipeline
 */
var PipelineMethods = /** @class */ (function () {
    function PipelineMethods() {
    }
    /**
     * Logs the start of the request
     */
    PipelineMethods.logStart = function (context) {
        return new Promise(function (resolve) {
            _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].log({
                data: _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].activeLogLevel === 1 /* Info */ ? {} : context,
                level: 1 /* Info */,
                message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Beginning " + context.method + " request (" + context.url + ")",
            });
            resolve(context);
        });
    };
    /**
     * Handles caching of the request
     */
    PipelineMethods.caching = function (context) {
        return new Promise(function (resolve) {
            // handle caching, if applicable
            if (context.useCaching) {
                _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Caching is enabled for request, checking cache...", 1 /* Info */);
                var cacheOptions = new _caching__WEBPACK_IMPORTED_MODULE_3__["CachingOptions"](context.url.toLowerCase());
                if (context.cachingOptions !== undefined) {
                    cacheOptions = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(cacheOptions, context.cachingOptions);
                }
                // we may not have a valid store
                if (cacheOptions.store !== null) {
                    // check if we have the data in cache and if so resolve the promise and return
                    var data = cacheOptions.store.get(cacheOptions.key);
                    if (data !== null) {
                        _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].log({
                            data: _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].activeLogLevel === 1 /* Info */ ? {} : data,
                            level: 1 /* Info */,
                            message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Value returned from cache.",
                        });
                        // ensure we clear any held batch dependency we are resolving from the cache
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isFunc"])(context.batchDependency)) {
                            context.batchDependency();
                        }
                        // handle the case where a parser needs to take special actions with a cached result
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(context.parser, "hydrate")) {
                            data = context.parser.hydrate(data);
                        }
                        return setResult(context, data).then(function (ctx) { return resolve(ctx); });
                    }
                }
                _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Value not found in cache.", 1 /* Info */);
                // if we don't then wrap the supplied parser in the caching parser wrapper
                // and send things on their way
                context.parser = new _caching__WEBPACK_IMPORTED_MODULE_3__["CachingParserWrapper"](context.parser, cacheOptions);
            }
            return resolve(context);
        });
    };
    /**
     * Sends the request
     */
    PipelineMethods.send = function (context) {
        return new Promise(function (resolve, reject) {
            // send or batch the request
            if (context.isBatched) {
                var p = context.batch.add(context);
                // we release the dependency here to ensure the batch does not execute until the request is added to the batch
                if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isFunc"])(context.batchDependency)) {
                    context.batchDependency();
                }
                _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Batching request in batch " + context.batch.batchId + ".", 1 /* Info */);
                // we set the result as the promise which will be resolved by the batch's execution
                resolve(setResult(context, p));
            }
            else {
                _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].write("[" + context.requestId + "] (" + (new Date()).getTime() + ") Sending request.", 1 /* Info */);
                // we are not part of a batch, so proceed as normal
                var client = context.clientFactory();
                var opts = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(context.options || {}, { method: context.method });
                client.fetch(context.url, opts)
                    .then(function (response) { return context.parser.parse(response); })
                    .then(function (result) { return setResult(context, result); })
                    .then(function (ctx) { return resolve(ctx); })
                    .catch(function (e) { return reject(e); });
            }
        });
    };
    /**
     * Logs the end of the request
     */
    PipelineMethods.logEnd = function (context) {
        return new Promise(function (resolve) {
            if (context.isBatched) {
                _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].log({
                    data: _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].activeLogLevel === 1 /* Info */ ? {} : context,
                    level: 1 /* Info */,
                    message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") " + context.method + " request will complete in batch " + context.batch.batchId + ".",
                });
            }
            else {
                _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].log({
                    data: _pnp_logging__WEBPACK_IMPORTED_MODULE_2__["Logger"].activeLogLevel === 1 /* Info */ ? {} : context,
                    level: 1 /* Info */,
                    message: "[" + context.requestId + "] (" + (new Date()).getTime() + ") Completing " + context.method + " request.",
                });
            }
            resolve(context);
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        requestPipelineMethod(true)
    ], PipelineMethods, "logStart", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        requestPipelineMethod()
    ], PipelineMethods, "caching", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        requestPipelineMethod()
    ], PipelineMethods, "send", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        requestPipelineMethod(true)
    ], PipelineMethods, "logEnd", null);
    return PipelineMethods;
}());

function getDefaultPipeline() {
    return [
        PipelineMethods.logStart,
        PipelineMethods.caching,
        PipelineMethods.send,
        PipelineMethods.logEnd,
    ].slice(0);
}
//# sourceMappingURL=pipeline.js.map

/***/ }),

/***/ "UK2s":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/operations.js ***!
  \********************************************/
/*! exports provided: spGet, spPost, spDelete, spPatch, spPostDelete, spPostDeleteETag */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spGet", function() { return spGet; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spPost", function() { return spPost; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spDelete", function() { return spDelete; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spPatch", function() { return spPatch; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spPostDelete", function() { return spPostDelete; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "spPostDeleteETag", function() { return spPostDeleteETag; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sphttpclient__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./sphttpclient */ "DqvY");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_toabsoluteurl__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./utils/toabsoluteurl */ "G2Ar");





var spClientBinder = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["defaultPipelineBinder"])(function () { return new _sphttpclient__WEBPACK_IMPORTED_MODULE_2__["SPHttpClient"](); });
var send = function (method) {
    var operation = spClientBinder(method);
    return function (o, options) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data, batchDependency, url;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        data = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["cloneQueryableData"])(o.data);
                        batchDependency = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["objectDefinedNotNull"])(data.batch) ? data.batch.addDependency() : function () { return; };
                        return [4 /*yield*/, Object(_utils_toabsoluteurl__WEBPACK_IMPORTED_MODULE_4__["toAbsoluteUrl"])(o.toUrlAndQuery())];
                    case 1:
                        url = _a.sent();
                        Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["mergeOptions"])(data.options, options);
                        return [2 /*return*/, operation(Object.assign({}, data, {
                                batchDependency: batchDependency,
                                url: url,
                            }))];
                }
            });
        });
    };
};
var spGet = function (o, options) {
    // Fix for #304 - when we clone objects we in some cases then execute a get request
    // in these cases the caching settings were getting dropped from the request
    // this tracks if the object from which this was cloned was caching and applies that to an immediate get request
    // does not affect objects cloned from this as we are using different fields to track the settings so it won't
    // be triggered
    if (o.data.cloneParentWasCaching) {
        o.usingCaching(o.data.cloneParentCacheOptions);
    }
    // if we are forcing caching set that in the data here
    if (o._forceCaching) {
        o.data.useCaching = true;
    }
    return send("GET")(o, options);
};
var spPost = send("POST");
var spDelete = send("DELETE");
var spPatch = send("PATCH");
var spPostDelete = function (o, options) {
    var opts = Object.assign(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["headers"])({ "X-HTTP-Method": "DELETE" }), options);
    return spPost(o, opts);
};
var spPostDeleteETag = function (o, options, eTag) {
    if (eTag === void 0) { eTag = "*"; }
    var opts = Object.assign(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["headers"])({ "X-HTTP-Method": "DELETE", "IF-Match": eTag }), options);
    return spPost(o, opts);
};
//# sourceMappingURL=operations.js.map

/***/ }),

/***/ "UKGb":
/*!***************************************!*\
  !*** ./node_modules/@pnp/sp/index.js ***!
  \***************************************/
/*! exports provided: SharePointQueryableInstance, SharePointQueryableCollection, SharePointQueryable, spInvokableFactory, SPBatch, SPHttpClient, SPRest, sp, toAbsoluteUrl, extractWebUrl, objectToSPKeyValueCollection, stripInvalidFileFolderChars, containsInvalidFileFolderChars, defaultPath, spGet, spPost, spDelete, spPatch, spPostDelete, spPostDeleteETag, PrincipalType, PrincipalSource, PageType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./sharepointqueryable */ "E6kk");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryableInstance", function() { return _sharepointqueryable__WEBPACK_IMPORTED_MODULE_0__["SharePointQueryableInstance"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryableCollection", function() { return _sharepointqueryable__WEBPACK_IMPORTED_MODULE_0__["SharePointQueryableCollection"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharePointQueryable", function() { return _sharepointqueryable__WEBPACK_IMPORTED_MODULE_0__["SharePointQueryable"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spInvokableFactory", function() { return _sharepointqueryable__WEBPACK_IMPORTED_MODULE_0__["spInvokableFactory"]; });

/* harmony import */ var _batch__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./batch */ "8xxP");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPBatch", function() { return _batch__WEBPACK_IMPORTED_MODULE_1__["SPBatch"]; });

/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./decorators */ "hMpi");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "defaultPath", function() { return _decorators__WEBPACK_IMPORTED_MODULE_2__["defaultPath"]; });

/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./operations */ "UK2s");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spGet", function() { return _operations__WEBPACK_IMPORTED_MODULE_3__["spGet"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spPost", function() { return _operations__WEBPACK_IMPORTED_MODULE_3__["spPost"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spDelete", function() { return _operations__WEBPACK_IMPORTED_MODULE_3__["spDelete"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spPatch", function() { return _operations__WEBPACK_IMPORTED_MODULE_3__["spPatch"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spPostDelete", function() { return _operations__WEBPACK_IMPORTED_MODULE_3__["spPostDelete"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "spPostDeleteETag", function() { return _operations__WEBPACK_IMPORTED_MODULE_3__["spPostDeleteETag"]; });

/* harmony import */ var _sphttpclient__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./sphttpclient */ "DqvY");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPHttpClient", function() { return _sphttpclient__WEBPACK_IMPORTED_MODULE_4__["SPHttpClient"]; });

/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./rest */ "9l4K");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPRest", function() { return _rest__WEBPACK_IMPORTED_MODULE_5__["SPRest"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "sp", function() { return _rest__WEBPACK_IMPORTED_MODULE_5__["sp"]; });

/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./types */ "tCQJ");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PrincipalType", function() { return _types__WEBPACK_IMPORTED_MODULE_6__["PrincipalType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PrincipalSource", function() { return _types__WEBPACK_IMPORTED_MODULE_6__["PrincipalSource"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PageType", function() { return _types__WEBPACK_IMPORTED_MODULE_6__["PageType"]; });

/* harmony import */ var _utils_toabsoluteurl__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./utils/toabsoluteurl */ "G2Ar");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "toAbsoluteUrl", function() { return _utils_toabsoluteurl__WEBPACK_IMPORTED_MODULE_7__["toAbsoluteUrl"]; });

/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./utils/extractweburl */ "tSEd");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "extractWebUrl", function() { return _utils_extractweburl__WEBPACK_IMPORTED_MODULE_8__["extractWebUrl"]; });

/* harmony import */ var _utils_objectToSPKeyValueCollection__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./utils/objectToSPKeyValueCollection */ "EVeO");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "objectToSPKeyValueCollection", function() { return _utils_objectToSPKeyValueCollection__WEBPACK_IMPORTED_MODULE_9__["objectToSPKeyValueCollection"]; });

/* harmony import */ var _utils_file_names__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ./utils/file-names */ "YFzv");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "stripInvalidFileFolderChars", function() { return _utils_file_names__WEBPACK_IMPORTED_MODULE_10__["stripInvalidFileFolderChars"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "containsInvalidFileFolderChars", function() { return _utils_file_names__WEBPACK_IMPORTED_MODULE_10__["containsInvalidFileFolderChars"]; });












//# sourceMappingURL=index.js.map

/***/ }),

/***/ "UPqE":
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/clientside-pages/types.js ***!
  \********************************************************/
/*! exports provided: PromotedState, _ClientsidePage, ClientsidePageFromFile, CreateClientsidePage, CanvasSection, CanvasColumn, ColumnControl, ClientsideText, ClientsideWebpart */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PromotedState", function() { return PromotedState; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_ClientsidePage", function() { return _ClientsidePage; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ClientsidePageFromFile", function() { return ClientsidePageFromFile; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CreateClientsidePage", function() { return CreateClientsidePage; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CanvasSection", function() { return CanvasSection; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CanvasColumn", function() { return CanvasColumn; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ColumnControl", function() { return ColumnControl; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ClientsideText", function() { return ClientsideText; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ClientsideWebpart", function() { return ClientsideWebpart; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _sites_types__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../sites/types */ "ilAS");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ./funcs */ "ld+d");
/* harmony import */ var _files_web__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ../files/web */ "SPHc");
/* harmony import */ var _comments_item__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! ../comments/item */ "NxAp");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! ../telemetry */ "h3gW");
















/**
 * Page promotion state
 */
var PromotedState;
(function (PromotedState) {
    /**
     * Regular client side page
     */
    PromotedState[PromotedState["NotPromoted"] = 0] = "NotPromoted";
    /**
     * Page that will be promoted as news article after publishing
     */
    PromotedState[PromotedState["PromoteOnPublish"] = 1] = "PromoteOnPublish";
    /**
     * Page that is promoted as news article
     */
    PromotedState[PromotedState["Promoted"] = 2] = "Promoted";
})(PromotedState || (PromotedState = {}));
function initFrom(o, url) {
    return ClientsidePage(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(o.toUrl()), url).configureFrom(o);
}
/**
 * Represents the data and methods associated with client side "modern" pages
 */
var _ClientsidePage = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_ClientsidePage, _super);
    /**
     * PLEASE DON'T USE THIS CONSTRUCTOR DIRECTLY, thank you 🐇
     */
    function _ClientsidePage(baseUrl, path, json, noInit, sections, commentsDisabled) {
        if (noInit === void 0) { noInit = false; }
        if (sections === void 0) { sections = []; }
        if (commentsDisabled === void 0) { commentsDisabled = false; }
        var _this = _super.call(this, baseUrl, path) || this;
        _this.json = json;
        _this.sections = sections;
        _this.commentsDisabled = commentsDisabled;
        _this._bannerImageDirty = false;
        // ensure we have a good url to build on for the pages api
        if (typeof baseUrl === "string") {
            _this.data.parentUrl = "";
            _this.data.url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["combine"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(baseUrl), path);
        }
        else {
            _this.assign(initFrom(baseUrl, null), path);
        }
        // set a default page settings slice
        _this._pageSettings = { controlType: 0, pageSettingsSlice: { isDefaultDescription: true, isDefaultThumbnail: true } };
        // set a default layout part
        _this._layoutPart = _ClientsidePage.getDefaultLayoutPart();
        if (typeof json !== "undefined" && !noInit) {
            _this.fromJSON(json);
        }
        return _this;
    }
    _ClientsidePage.getDefaultLayoutPart = function () {
        return {
            dataVersion: "1.4",
            description: "Title Region Description",
            id: "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
            instanceId: "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
            properties: {
                authorByline: [],
                authors: [],
                layoutType: "FullWidthImage",
                showPublishDate: false,
                showTopicHeader: false,
                textAlignment: "Left",
                title: "",
                topicHeader: "",
            },
            serverProcessedContent: { htmlStrings: {}, searchablePlainTexts: {}, imageSources: {}, links: {} },
            title: "Title area",
        };
    };
    Object.defineProperty(_ClientsidePage.prototype, "pageLayout", {
        get: function () {
            return this.json.PageLayoutType;
        },
        set: function (value) {
            this.json.PageLayoutType = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "bannerImageUrl", {
        get: function () {
            return this.json.BannerImageUrl;
        },
        set: function (value) {
            this.json.BannerImageUrl = value;
            this._bannerImageDirty = true;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "topicHeader", {
        get: function () {
            return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["objectDefinedNotNull"])(this.json.TopicHeader) ? this.json.TopicHeader : "";
        },
        set: function (value) {
            this.json.TopicHeader = value;
            this._layoutPart.properties.topicHeader = value;
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["stringIsNullOrEmpty"])(value)) {
                this.showTopicHeader = false;
            }
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "title", {
        get: function () {
            return this._layoutPart.properties.title;
        },
        set: function (value) {
            this.json.Title = value;
            this._layoutPart.properties.title = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "layoutType", {
        get: function () {
            return this._layoutPart.properties.layoutType;
        },
        set: function (value) {
            this._layoutPart.properties.layoutType = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "headerTextAlignment", {
        get: function () {
            return this._layoutPart.properties.textAlignment;
        },
        set: function (value) {
            this._layoutPart.properties.textAlignment = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "showTopicHeader", {
        get: function () {
            return this._layoutPart.properties.showTopicHeader;
        },
        set: function (value) {
            this._layoutPart.properties.showTopicHeader = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "showPublishDate", {
        get: function () {
            return this._layoutPart.properties.showPublishDate;
        },
        set: function (value) {
            this._layoutPart.properties.showPublishDate = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "hasVerticalSection", {
        get: function () {
            return this.sections.findIndex(function (s) { return s.layoutIndex === 2; }) > -1;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "authorByLine", {
        get: function () {
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["isArray"])(this.json.AuthorByline) && this.json.AuthorByline.length > 0) {
                return this.json.AuthorByline[0];
            }
            return null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_ClientsidePage.prototype, "verticalSection", {
        get: function () {
            if (this.hasVerticalSection) {
                return this.addVerticalSection();
            }
            return null;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Add a section to this page
     */
    _ClientsidePage.prototype.addSection = function () {
        var section = new CanvasSection(this, Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["getNextOrder"])(this.sections), 1);
        this.sections.push(section);
        return section;
    };
    /**
     * Add a section to this page
     */
    _ClientsidePage.prototype.addVerticalSection = function () {
        // we can only have one vertical section so we find it if it exists
        var sectionIndex = this.sections.findIndex(function (s) { return s.layoutIndex === 2; });
        if (sectionIndex > -1) {
            return this.sections[sectionIndex];
        }
        var section = new CanvasSection(this, Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["getNextOrder"])(this.sections), 2);
        this.sections.push(section);
        return section;
    };
    /**
     * Loads this instance from the appropriate JSON data
     *
     * @param pageData JSON data to load (replaces any existing data)
     */
    _ClientsidePage.prototype.fromJSON = function (pageData) {
        this.json = pageData;
        var canvasControls = JSON.parse(pageData.CanvasContent1);
        var layouts = JSON.parse(pageData.LayoutWebpartsContent);
        if (layouts && layouts.length > 0) {
            this._layoutPart = layouts[0];
        }
        this.setControls(canvasControls);
        return this;
    };
    /**
     * Loads this page's content from the server
     */
    _ClientsidePage.prototype.load = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var item, pageData;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.getItem("Id", "CommentsDisabled")];
                    case 1:
                        item = _a.sent();
                        return [4 /*yield*/, Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_4__["SharePointQueryable"])(this, "_api/sitepages/pages(" + item.Id + ")")()];
                    case 2:
                        pageData = _a.sent();
                        this.commentsDisabled = item.CommentsDisabled;
                        return [2 /*return*/, this.fromJSON(pageData)];
                }
            });
        });
    };
    /**
     * Persists the content changes (sections, columns, and controls) [does not work with batching]
     *
     * @param publish If true the page is published, if false the changes are persisted to SharePoint but not published [Default: true]
     */
    _ClientsidePage.prototype.save = function (publish) {
        if (publish === void 0) { publish = true; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var origImgUrl, site, web, imgFile, siteId_1, webId_1, imgId_1, listId_1, webUrl_1, batch, f, saveBody, updater, r;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.json.Id === null) {
                            throw Error("The id for this page is null. If you want to create a new page, please use ClientSidePage.Create");
                        }
                        if (!this._bannerImageDirty) return [3 /*break*/, 2];
                        origImgUrl = this.json.BannerImageUrl;
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["isUrlAbsolute"])(origImgUrl)) {
                            // do our best to make this a server relative url by removing the x.sharepoint.com part
                            origImgUrl = origImgUrl.replace(/^https?:\/\/[a-z0-9\.]*?\.[a-z]{2,3}\//i, "/");
                        }
                        site = Object(_sites_types__WEBPACK_IMPORTED_MODULE_10__["Site"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(this.toUrl()));
                        web = Object(_webs_types__WEBPACK_IMPORTED_MODULE_8__["Web"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(this.toUrl()));
                        imgFile = web.getFileByServerRelativePath(origImgUrl);
                        siteId_1 = "";
                        webId_1 = "";
                        imgId_1 = "";
                        listId_1 = "";
                        webUrl_1 = "";
                        batch = web.createBatch();
                        site.select("Id", "Url").inBatch(batch)().then(function (r1) { return siteId_1 = r1.Id; });
                        web.select("Id", "Url").inBatch(batch)().then(function (r2) { webId_1 = r2.Id; webUrl_1 = r2.Url; });
                        imgFile.listItemAllFields.select("UniqueId", "ParentList/Id").expand("ParentList").inBatch(batch)()
                            .then(function (r3) { imgId_1 = r3.UniqueId; listId_1 = r3.ParentList.Id; });
                        // we know the .then calls above will run before execute resolves, ensuring the vars are set
                        return [4 /*yield*/, batch.execute()];
                    case 1:
                        // we know the .then calls above will run before execute resolves, ensuring the vars are set
                        _a.sent();
                        f = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_4__["SharePointQueryable"])(webUrl_1, "_layouts/15/getpreview.ashx");
                        f.query.set("guidSite", "" + siteId_1);
                        f.query.set("guidWeb", "" + webId_1);
                        f.query.set("guidFile", "" + imgId_1);
                        this.bannerImageUrl = f.toUrlAndQuery();
                        if (!Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["objectDefinedNotNull"])(this._layoutPart.serverProcessedContent)) {
                            this._layoutPart.serverProcessedContent = {};
                        }
                        this._layoutPart.serverProcessedContent.imageSources = { imageSource: origImgUrl };
                        if (!Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["objectDefinedNotNull"])(this._layoutPart.serverProcessedContent.customMetadata)) {
                            this._layoutPart.serverProcessedContent.customMetadata = {};
                        }
                        this._layoutPart.serverProcessedContent.customMetadata.imageSource = {
                            listId: listId_1,
                            siteId: siteId_1,
                            uniqueId: imgId_1,
                            webId: webId_1,
                        };
                        this._layoutPart.properties.webId = webId_1;
                        this._layoutPart.properties.siteId = siteId_1;
                        this._layoutPart.properties.listId = listId_1;
                        this._layoutPart.properties.uniqueId = imgId_1;
                        _a.label = 2;
                    case 2:
                        if (!!this.json.IsPageCheckedOutToCurrentUser) return [3 /*break*/, 4];
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_11__["spPost"])(initFrom(this, "_api/sitepages/pages(" + this.json.Id + ")/checkoutpage"))];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4:
                        saveBody = Object.assign(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.Publishing.SitePage"), {
                            AuthorByline: this.json.AuthorByline || [],
                            BannerImageUrl: this.bannerImageUrl,
                            CanvasContent1: this.getCanvasContent1(),
                            LayoutWebpartsContent: this.getLayoutWebpartsContent(),
                            Title: this.title,
                            TopicHeader: this.topicHeader,
                        });
                        updater = initFrom(this, "_api/sitepages/pages(" + this.json.Id + ")/savepage");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_11__["spPost"])(updater, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["headers"])({ "if-match": "*" }, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(saveBody)))];
                    case 5:
                        _a.sent();
                        r = true;
                        if (!publish) return [3 /*break*/, 7];
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_11__["spPost"])(initFrom(this, "_api/sitepages/pages(" + this.json.Id + ")/publish"))];
                    case 6:
                        r = _a.sent();
                        if (r) {
                            this.json.IsPageCheckedOutToCurrentUser = false;
                        }
                        _a.label = 7;
                    case 7:
                        this._bannerImageDirty = false;
                        return [2 /*return*/, r];
                }
            });
        });
    };
    /**
     * Discards the checkout of this page
     */
    _ClientsidePage.prototype.discardPageCheckout = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var d;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.json.Id === null) {
                            throw Error("The id for this page is null. If you want to create a new page, please use ClientSidePage.Create");
                        }
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_11__["spPost"])(initFrom(this, "_api/sitepages/pages(" + this.json.Id + ")/discardPage"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.Publishing.SitePage")))];
                    case 1:
                        d = _a.sent();
                        this.fromJSON(d);
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Promotes this page as a news item
     */
    _ClientsidePage.prototype.promoteToNews = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                return [2 /*return*/, this.promoteNewsImpl("promoteToNews")];
            });
        });
    };
    // API is currently broken on server side
    // public async demoteFromNews(): Promise<boolean> {
    //     return this.promoteNewsImpl("demoteFromNews");
    // }
    /**
     * Finds a control by the specified instance id
     *
     * @param id Instance id of the control to find
     */
    _ClientsidePage.prototype.findControlById = function (id) {
        return this.findControl(function (c) { return c.id === id; });
    };
    /**
     * Finds a control within this page's control tree using the supplied predicate
     *
     * @param predicate Takes a control and returns true or false, if true that control is returned by findControl
     */
    _ClientsidePage.prototype.findControl = function (predicate) {
        // check all sections
        for (var i = 0; i < this.sections.length; i++) {
            // check all columns
            for (var j = 0; j < this.sections[i].columns.length; j++) {
                // check all controls
                for (var k = 0; k < this.sections[i].columns[j].controls.length; k++) {
                    // check to see if the predicate likes this control
                    if (predicate(this.sections[i].columns[j].controls[k])) {
                        return this.sections[i].columns[j].controls[k];
                    }
                }
            }
        }
        // we found nothing so give nothing back
        return null;
    };
    /**
     * Creates a copy of this page
     *
     * @param web The web where we will create the copy
     * @param pageName The file name of the new page
     * @param title The title of the new page
     * @param publish If true the page will be published
     */
    _ClientsidePage.prototype.copy = function (web, pageName, title, publish, promotedState) {
        if (publish === void 0) { publish = true; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var page, url, makeGuid, guidSite, guidWeb, guidFile, site, id, openWeb, file, props;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, CreateClientsidePage(web, pageName, title, this.pageLayout, promotedState)];
                    case 1:
                        page = _a.sent();
                        // we know the method is on the class - but it is protected so not part of the interface
                        page.setControls(this.getControls());
                        if (!!Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["stringIsNullOrEmpty"])(this.json.BannerImageUrl)) return [3 /*break*/, 5];
                        url = new URL(this.json.BannerImageUrl);
                        makeGuid = function (s) { return s.replace(/(.{8})(.{4})(.{4})(.{4})(.{12})/g, "$1-$2-$3-$4-$5"); };
                        if (!(url.searchParams.has("guidSite") && url.searchParams.has("guidWeb") && url.searchParams.has("guidFile"))) return [3 /*break*/, 5];
                        guidSite = makeGuid(url.searchParams.get("guidSite"));
                        guidWeb = makeGuid(url.searchParams.get("guidWeb"));
                        guidFile = makeGuid(url.searchParams.get("guidFile"));
                        site = Object(_sites_types__WEBPACK_IMPORTED_MODULE_10__["Site"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(this.toUrl()));
                        return [4 /*yield*/, site.select("Id")()];
                    case 2:
                        id = _a.sent();
                        if (!(id.Id === guidSite)) return [3 /*break*/, 5];
                        return [4 /*yield*/, site.openWebById(guidWeb)];
                    case 3:
                        openWeb = _a.sent();
                        return [4 /*yield*/, openWeb.web.getFileById(guidFile).select("ServerRelativeUrl")()];
                    case 4:
                        file = _a.sent();
                        props = {};
                        if (this._layoutPart.properties) {
                            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(this._layoutPart.properties, "translateX")) {
                                props.translateX = this._layoutPart.properties.translateX;
                            }
                            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(this._layoutPart.properties, "translateY")) {
                                props.translateY = this._layoutPart.properties.translateY;
                            }
                            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(this._layoutPart.properties, "imageSourceType")) {
                                props.imageSourceType = this._layoutPart.properties.imageSourceType;
                            }
                            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(this._layoutPart.properties, "altText")) {
                                props.altText = this._layoutPart.properties.altText;
                            }
                        }
                        page.setBannerImage(file.ServerRelativeUrl, props);
                        _a.label = 5;
                    case 5: return [4 /*yield*/, page.save(publish)];
                    case 6:
                        _a.sent();
                        return [2 /*return*/, page];
                }
            });
        });
    };
    /**
     * Sets the modern page banner image
     *
     * @param url Url of the image to display
     * @param altText Alt text to describe the image
     * @param bannerProps Additional properties to control display of the banner
     */
    _ClientsidePage.prototype.setBannerImage = function (url, props) {
        this.bannerImageUrl = url;
        this._layoutPart.properties.imageSourceType = 2; // this seems to always be true, so default?
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["objectDefinedNotNull"])(props)) {
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(props, "translateX")) {
                this._layoutPart.properties.translateX = props.translateX;
            }
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(props, "translateY")) {
                this._layoutPart.properties.translateY = props.translateY;
            }
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(props, "imageSourceType")) {
                this._layoutPart.properties.imageSourceType = props.imageSourceType;
            }
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(props, "altText")) {
                this._layoutPart.properties.altText = props.altText;
            }
        }
    };
    /**
     * Sets the authors for this page from the supplied list of user integer ids
     *
     * @param authorId The integer id of the user to set as the author
     */
    _ClientsidePage.prototype.setAuthorById = function (authorId) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var userLoginData;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_4__["SharePointQueryableCollection"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(this.toUrl()), "/_api/web/siteusers")
                            .configureFrom(this)
                            .filter("Id eq " + authorId)
                            .select("LoginName")()];
                    case 1:
                        userLoginData = _a.sent();
                        if (userLoginData.length < 1) {
                            throw Error("Could not find user with id " + authorId + ".");
                        }
                        return [2 /*return*/, this.setAuthorByLoginName(userLoginData[0].LoginName)];
                }
            });
        });
    };
    /**
     * Sets the authors for this page from the supplied list of user integer ids
     *
     * @param authorLoginName The login name of the user (ex: i:0#.f|membership|name@tenant.com)
     */
    _ClientsidePage.prototype.setAuthorByLoginName = function (authorLoginName) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var userLoginData;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_4__["SharePointQueryableCollection"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(this.toUrl()), "/_api/web/siteusers")
                            .configureFrom(this)
                            .filter("LoginName eq '" + encodeURIComponent(authorLoginName) + "'")
                            .select("UserPrincipalName", "Title")()];
                    case 1:
                        userLoginData = _a.sent();
                        if (userLoginData.length < 1) {
                            throw Error("Could not find user with login name '" + authorLoginName + "'.");
                        }
                        this.json.AuthorByline = [authorLoginName];
                        this._layoutPart.properties.authorByline = [authorLoginName];
                        this._layoutPart.properties.authors = [{
                                id: authorLoginName,
                                name: userLoginData[0].Title,
                                role: "",
                                upn: userLoginData[0].UserPrincipalName,
                            }];
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Gets the list item associated with this clientside page
     *
     * @param selects Specific set of fields to include when getting the item
     */
    _ClientsidePage.prototype.getItem = function () {
        var selects = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            selects[_i] = arguments[_i];
        }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var initer, listData, item, itemData;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        initer = initFrom(this, "/_api/lists/EnsureClientRenderedSitePagesLibrary").select("EnableModeration", "EnableMinorVersions", "Id");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_11__["spPost"])(initer)];
                    case 1:
                        listData = _a.sent();
                        item = (Object(_lists_types__WEBPACK_IMPORTED_MODULE_6__["List"])(listData["odata.id"])).configureFrom(this).items.getById(this.json.Id);
                        return [4 /*yield*/, item.select.apply(item, selects)()];
                    case 2:
                        itemData = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])((Object(_items_types__WEBPACK_IMPORTED_MODULE_3__["Item"])(Object(_odata__WEBPACK_IMPORTED_MODULE_7__["odataUrlFrom"])(itemData))).configureFrom(this), itemData)];
                }
            });
        });
    };
    /**
     * Extends this queryable from the provided parent
     *
     * @param parent Parent queryable from which we will derive a base url
     * @param path Additional path
     */
    _ClientsidePage.prototype.assign = function (parent, path) {
        this.data.parentUrl = parent.data.url;
        this.data.url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["combine"])(this.data.parentUrl, path || "");
        this.configureFrom(parent);
    };
    _ClientsidePage.prototype.getCanvasContent1 = function () {
        return JSON.stringify(this.getControls());
    };
    _ClientsidePage.prototype.getLayoutWebpartsContent = function () {
        if (this._layoutPart) {
            return JSON.stringify([this._layoutPart]);
        }
        else {
            return JSON.stringify(null);
        }
    };
    _ClientsidePage.prototype.setControls = function (controls) {
        if (controls && controls.length) {
            for (var i = 0; i < controls.length; i++) {
                // if no control type is present this is a column which we give type 0 to let us process it
                var controlType = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(controls[i], "controlType") ? controls[i].controlType : 0;
                switch (controlType) {
                    case 0:
                        // empty canvas column or page settings
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(controls[i], "pageSettingsSlice")) {
                            this._pageSettings = controls[i];
                        }
                        else {
                            // we have an empty column
                            this.mergeColumnToTree(new CanvasColumn(controls[i]));
                        }
                        break;
                    case 3:
                        var part = new ClientsideWebpart(controls[i]);
                        this.mergePartToTree(part, part.data.position);
                        break;
                    case 4:
                        var textData = controls[i];
                        var text = new ClientsideText(textData.innerHTML, textData);
                        this.mergePartToTree(text, text.data.position);
                        break;
                }
            }
            Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["reindex"])(this.sections);
        }
    };
    _ClientsidePage.prototype.getControls = function () {
        var _this = this;
        // reindex things
        Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["reindex"])(this.sections);
        // rollup the control changes
        var canvasData = [];
        this.sections.forEach(function (section) {
            section.columns.forEach(function (column) {
                if (column.controls.length < 1) {
                    // empty column
                    canvasData.push({
                        displayMode: column.data.displayMode,
                        emphasis: _this.getEmphasisObj(section.emphasis),
                        position: column.data.position,
                    });
                }
                else {
                    column.controls.forEach(function (control) {
                        control.data.emphasis = _this.getEmphasisObj(section.emphasis);
                        canvasData.push(control.data);
                    });
                }
            });
        });
        canvasData.push(this._pageSettings);
        return canvasData;
    };
    _ClientsidePage.prototype.getEmphasisObj = function (value) {
        if (value < 1 || value > 3) {
            return {};
        }
        return { zoneEmphasis: value };
    };
    _ClientsidePage.prototype.promoteNewsImpl = function (method) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var lastPubData;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.json.Id === null) {
                            throw Error("The id for this page is null.");
                        }
                        if (!Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["stringIsNullOrEmpty"])(this.json.VersionInfo.LastVersionCreatedBy)) return [3 /*break*/, 2];
                        lastPubData = new Date(this.json.VersionInfo.LastVersionCreated);
                        if (!(lastPubData.getFullYear() < 2000)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.save(true)];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_11__["spPost"])(initFrom(this, "_api/sitepages/pages(" + this.json.Id + ")/" + method), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.Publishing.SitePage")))];
                    case 3: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    /**
     * Merges the control into the tree of sections and columns for this page
     *
     * @param control The control to merge
     */
    _ClientsidePage.prototype.mergePartToTree = function (control, positionData) {
        var _a, _b, _c;
        var column = null;
        var sectionFactor = 12;
        var sectionIndex = 0;
        var zoneIndex = 0;
        var layoutIndex = 1;
        // handle case where we don't have position data (shouldn't happen?)
        if (positionData) {
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(positionData, "zoneIndex")) {
                zoneIndex = positionData.zoneIndex;
            }
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(positionData, "sectionIndex")) {
                sectionIndex = positionData.sectionIndex;
            }
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(positionData, "sectionFactor")) {
                sectionFactor = positionData.sectionFactor;
            }
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(positionData, "layoutIndex")) {
                layoutIndex = positionData.layoutIndex;
            }
        }
        var zoneEmphasis = (_c = (_b = (_a = control.data) === null || _a === void 0 ? void 0 : _a.emphasis) === null || _b === void 0 ? void 0 : _b.zoneEmphasis, (_c !== null && _c !== void 0 ? _c : 0));
        var section = this.getOrCreateSection(zoneIndex, layoutIndex, zoneEmphasis);
        var columns = section.columns.filter(function (c) { return c.order === sectionIndex; });
        if (columns.length < 1) {
            column = section.addColumn(sectionFactor, layoutIndex);
        }
        else {
            column = columns[0];
        }
        control.column = column;
        column.addControl(control);
    };
    /**
     * Merges the supplied column into the tree
     *
     * @param column Column to merge
     * @param position The position data for the column
     */
    _ClientsidePage.prototype.mergeColumnToTree = function (column) {
        var order = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(column.data, "position") && Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(column.data.position, "zoneIndex") ? column.data.position.zoneIndex : 0;
        var layoutIndex = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(column.data, "position") && Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(column.data.position, "layoutIndex") ? column.data.position.layoutIndex : 1;
        var section = this.getOrCreateSection(order, layoutIndex, column.data.emphasis.zoneEmphasis || 0);
        column.section = section;
        section.columns.push(column);
    };
    /**
     * Handle the logic to get or create a section based on the supplied order and layoutIndex
     *
     * @param order Section order
     * @param layoutIndex Layout Index (1 === normal, 2 === vertical section)
     * @param emphasis The section emphasis
     */
    _ClientsidePage.prototype.getOrCreateSection = function (order, layoutIndex, emphasis) {
        var section = null;
        var sections = this.sections.filter(function (s) { return s.order === order && s.layoutIndex === layoutIndex; });
        if (sections.length < 1) {
            section = layoutIndex === 2 ? this.addVerticalSection() : this.addSection();
            section.order = order;
            section.emphasis = emphasis;
        }
        else {
            section = sections[0];
        }
        return section;
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_15__["tag"])("csp.load")
    ], _ClientsidePage.prototype, "load", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_15__["tag"])("csp.save")
    ], _ClientsidePage.prototype, "save", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_15__["tag"])("csp.discardPageCheckout")
    ], _ClientsidePage.prototype, "discardPageCheckout", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_15__["tag"])("csp.promoteToNews")
    ], _ClientsidePage.prototype, "promoteToNews", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_15__["tag"])("csp.copy")
    ], _ClientsidePage.prototype, "copy", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_15__["tag"])("csp.getItem")
    ], _ClientsidePage.prototype, "getItem", null);
    return _ClientsidePage;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_4__["_SharePointQueryable"]));

/**
 * Invokable factory for IClientSidePage instances
 */
var ClientsidePage = function (baseUrl, path, json, noInit, sections, commentsDisabled) {
    if (noInit === void 0) { noInit = false; }
    if (sections === void 0) { sections = []; }
    if (commentsDisabled === void 0) { commentsDisabled = false; }
    return Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["invokableFactory"])(_ClientsidePage)(baseUrl, path, json, noInit, sections, commentsDisabled);
};
/**
 * Loads a client side page from the supplied IFile instance
 *
 * @param file Source IFile instance
 */
var ClientsidePageFromFile = function (file) { return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(void 0, void 0, void 0, function () {
    var item, page;
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
        switch (_a.label) {
            case 0: return [4 /*yield*/, file.getItem()];
            case 1:
                item = _a.sent();
                page = ClientsidePage(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(file.toUrl()), "", { Id: item.Id }, true);
                return [2 /*return*/, page.configureFrom(file).load()];
        }
    });
}); };
/**
 * Creates a new client side page
 *
 * @param web The web or list
 * @param pageName The name of the page (filename)
 * @param title The page's title
 * @param PageLayoutType Layout to use when creating the page
 */
var CreateClientsidePage = function (web, pageName, title, PageLayoutType, promotedState) {
    if (PageLayoutType === void 0) { PageLayoutType = "Article"; }
    if (promotedState === void 0) { promotedState = 0; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(void 0, void 0, void 0, function () {
        var pageInitData, newPage;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    // patched because previously we used the full page name with the .aspx at the end
                    // this allows folk's existing code to work after the re-write to the new API
                    pageName = pageName.replace(/\.aspx$/i, "");
                    return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_11__["spPost"])(initFrom(web, "_api/sitepages/pages"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(Object.assign(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.Publishing.SitePage"), {
                            PageLayoutType: PageLayoutType,
                            PromotedState: promotedState,
                        })))];
                case 1:
                    pageInitData = _a.sent();
                    newPage = ClientsidePage(web, "", pageInitData);
                    newPage.title = pageName;
                    return [4 /*yield*/, newPage.save(false)];
                case 2:
                    _a.sent();
                    newPage.title = title;
                    return [2 /*return*/, newPage];
            }
        });
    });
};
var CanvasSection = /** @class */ (function () {
    function CanvasSection(page, order, layoutIndex, columns, _emphasis) {
        if (columns === void 0) { columns = []; }
        if (_emphasis === void 0) { _emphasis = 0; }
        this.page = page;
        this.columns = columns;
        this._emphasis = _emphasis;
        this._memId = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["getGUID"])();
        this._order = order;
        this._layoutIndex = layoutIndex;
    }
    Object.defineProperty(CanvasSection.prototype, "order", {
        get: function () {
            return this._order;
        },
        set: function (value) {
            this._order = value;
            for (var i = 0; i < this.columns.length; i++) {
                this.columns[i].data.position.zoneIndex = value;
            }
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CanvasSection.prototype, "layoutIndex", {
        get: function () {
            return this._layoutIndex;
        },
        set: function (value) {
            this._layoutIndex = value;
            for (var i = 0; i < this.columns.length; i++) {
                this.columns[i].data.position.layoutIndex = value;
            }
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CanvasSection.prototype, "defaultColumn", {
        /**
         * Default column (this.columns[0]) for this section
         */
        get: function () {
            if (this.columns.length < 1) {
                this.addColumn(12);
            }
            return this.columns[0];
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Adds a new column to this section
     */
    CanvasSection.prototype.addColumn = function (factor, layoutIndex) {
        if (layoutIndex === void 0) { layoutIndex = 1; }
        var column = new CanvasColumn();
        column.section = this;
        column.data.position.zoneIndex = this.order;
        column.data.position.layoutIndex = layoutIndex;
        column.data.position.sectionFactor = factor;
        column.order = Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["getNextOrder"])(this.columns);
        this.columns.push(column);
        return column;
    };
    /**
     * Adds a control to the default column for this section
     *
     * @param control Control to add to the default column
     */
    CanvasSection.prototype.addControl = function (control) {
        this.defaultColumn.addControl(control);
        return this;
    };
    Object.defineProperty(CanvasSection.prototype, "emphasis", {
        get: function () {
            return this._emphasis;
        },
        set: function (value) {
            this._emphasis = value;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Removes this section and all contained columns and controls from the collection
     */
    CanvasSection.prototype.remove = function () {
        var _this = this;
        this.page.sections = this.page.sections.filter(function (section) { return section._memId !== _this._memId; });
        Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["reindex"])(this.page.sections);
    };
    return CanvasSection;
}());

var CanvasColumn = /** @class */ (function () {
    function CanvasColumn(json, controls) {
        if (json === void 0) { json = JSON.parse(JSON.stringify(CanvasColumn.Default)); }
        if (controls === void 0) { controls = []; }
        this.json = json;
        this.controls = controls;
        this._section = null;
        this._memId = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["getGUID"])();
    }
    Object.defineProperty(CanvasColumn.prototype, "data", {
        get: function () {
            return this.json;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CanvasColumn.prototype, "section", {
        get: function () {
            return this._section;
        },
        set: function (section) {
            this._section = section;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CanvasColumn.prototype, "order", {
        get: function () {
            return this.data.position.sectionIndex;
        },
        set: function (value) {
            this.data.position.sectionIndex = value;
            for (var i = 0; i < this.controls.length; i++) {
                this.controls[i].data.position.zoneIndex = this.data.position.zoneIndex;
                this.controls[i].data.position.layoutIndex = this.data.position.layoutIndex;
                this.controls[i].data.position.sectionIndex = value;
            }
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(CanvasColumn.prototype, "factor", {
        get: function () {
            return this.data.position.sectionFactor;
        },
        set: function (value) {
            this.data.position.sectionFactor = value;
        },
        enumerable: true,
        configurable: true
    });
    CanvasColumn.prototype.addControl = function (control) {
        control.column = this;
        this.controls.push(control);
        return this;
    };
    CanvasColumn.prototype.getControl = function (index) {
        return this.controls[index];
    };
    CanvasColumn.prototype.remove = function () {
        var _this = this;
        this.section.columns = this.section.columns.filter(function (column) { return column._memId !== _this._memId; });
        Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["reindex"])(this.section.columns);
    };
    CanvasColumn.Default = {
        controlType: 0,
        displayMode: 2,
        emphasis: {},
        position: {
            layoutIndex: 1,
            sectionFactor: 12,
            sectionIndex: 1,
            zoneIndex: 1,
        },
    };
    return CanvasColumn;
}());

var ColumnControl = /** @class */ (function () {
    function ColumnControl(json) {
        this.json = json;
    }
    Object.defineProperty(ColumnControl.prototype, "id", {
        get: function () {
            return this.json.id;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ColumnControl.prototype, "data", {
        get: function () {
            return this.json;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ColumnControl.prototype, "column", {
        get: function () {
            return this._column;
        },
        set: function (value) {
            this._column = value;
            this.onColumnChange(this._column);
        },
        enumerable: true,
        configurable: true
    });
    ColumnControl.prototype.remove = function () {
        var _this = this;
        this.column.controls = this.column.controls.filter(function (control) { return control.id !== _this.id; });
        Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["reindex"])(this.column.controls);
    };
    ColumnControl.prototype.setData = function (data) {
        this.json = data;
    };
    return ColumnControl;
}());

var ClientsideText = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(ClientsideText, _super);
    function ClientsideText(text, json) {
        if (json === void 0) { json = JSON.parse(JSON.stringify(ClientsideText.Default)); }
        var _this = this;
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["stringIsNullOrEmpty"])(json.id)) {
            json.id = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["getGUID"])();
            json.anchorComponentId = json.id;
        }
        _this = _super.call(this, json) || this;
        _this.text = text;
        return _this;
    }
    Object.defineProperty(ClientsideText.prototype, "text", {
        get: function () {
            return this.data.innerHTML;
        },
        set: function (value) {
            if (!value.startsWith("<p>")) {
                value = "<p>" + value + "</p>";
            }
            this.data.innerHTML = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientsideText.prototype, "order", {
        get: function () {
            return this.data.position.controlIndex;
        },
        set: function (value) {
            this.data.position.controlIndex = value;
        },
        enumerable: true,
        configurable: true
    });
    ClientsideText.prototype.onColumnChange = function (col) {
        this.data.position.sectionFactor = col.factor;
        this.data.position.controlIndex = Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["getNextOrder"])(col.controls);
        this.data.position.zoneIndex = col.data.position.zoneIndex;
        this.data.position.sectionIndex = col.order;
        this.data.position.layoutIndex = col.data.position.layoutIndex;
    };
    ClientsideText.Default = {
        addedFromPersistedData: false,
        anchorComponentId: "",
        controlType: 4,
        displayMode: 2,
        editorType: "CKEditor",
        emphasis: {},
        id: "",
        innerHTML: "",
        position: {
            controlIndex: 1,
            layoutIndex: 1,
            sectionFactor: 12,
            sectionIndex: 1,
            zoneIndex: 1,
        },
    };
    return ClientsideText;
}(ColumnControl));

var ClientsideWebpart = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(ClientsideWebpart, _super);
    function ClientsideWebpart(json) {
        if (json === void 0) { json = JSON.parse(JSON.stringify(ClientsideWebpart.Default)); }
        return _super.call(this, json) || this;
    }
    ClientsideWebpart.fromComponentDef = function (definition) {
        var part = new ClientsideWebpart();
        part.import(definition);
        return part;
    };
    Object.defineProperty(ClientsideWebpart.prototype, "title", {
        get: function () {
            return this.data.webPartData.title;
        },
        set: function (value) {
            this.data.webPartData.title = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientsideWebpart.prototype, "description", {
        get: function () {
            return this.data.webPartData.description;
        },
        set: function (value) {
            this.data.webPartData.description = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientsideWebpart.prototype, "order", {
        get: function () {
            return this.data.position.controlIndex;
        },
        set: function (value) {
            this.data.position.controlIndex = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientsideWebpart.prototype, "height", {
        get: function () {
            return this.data.reservedHeight;
        },
        set: function (value) {
            this.data.reservedHeight = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientsideWebpart.prototype, "width", {
        get: function () {
            return this.data.reservedWidth;
        },
        set: function (value) {
            this.data.reservedWidth = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ClientsideWebpart.prototype, "dataVersion", {
        get: function () {
            return this.data.webPartData.dataVersion;
        },
        set: function (value) {
            this.data.webPartData.dataVersion = value;
        },
        enumerable: true,
        configurable: true
    });
    ClientsideWebpart.prototype.setProperties = function (properties) {
        this.data.webPartData.properties = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["assign"])(this.data.webPartData.properties, properties);
        return this;
    };
    ClientsideWebpart.prototype.getProperties = function () {
        return this.data.webPartData.properties;
    };
    ClientsideWebpart.prototype.onColumnChange = function (col) {
        this.data.position.sectionFactor = col.factor;
        this.data.position.controlIndex = Object(_funcs__WEBPACK_IMPORTED_MODULE_12__["getNextOrder"])(col.controls);
        this.data.position.zoneIndex = col.data.position.zoneIndex;
        this.data.position.sectionIndex = col.data.position.sectionIndex;
        this.data.position.layoutIndex = col.data.position.layoutIndex;
    };
    ClientsideWebpart.prototype.import = function (component) {
        var id = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["getGUID"])();
        var componendId = component.Id.replace(/^\{|\}$/g, "").toLowerCase();
        var manifest = JSON.parse(component.Manifest);
        var preconfiguredEntries = manifest.preconfiguredEntries[0];
        this.setData(Object.assign({}, this.data, {
            id: id,
            webPartData: {
                dataVersion: "1.0",
                description: preconfiguredEntries.description.default,
                id: componendId,
                instanceId: id,
                properties: preconfiguredEntries.properties,
                title: preconfiguredEntries.title.default,
            },
            webPartId: componendId,
        }));
    };
    ClientsideWebpart.Default = {
        addedFromPersistedData: false,
        controlType: 3,
        displayMode: 2,
        emphasis: {},
        id: null,
        position: {
            controlIndex: 1,
            layoutIndex: 1,
            sectionFactor: 12,
            sectionIndex: 1,
            zoneIndex: 1,
        },
        reservedHeight: 500,
        reservedWidth: 500,
        webPartData: null,
        webPartId: null,
    };
    return ClientsideWebpart;
}(ColumnControl));

//# sourceMappingURL=types.js.map

/***/ }),

/***/ "UWqr":
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_UWqr__;

/***/ }),

/***/ "UuEs":
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/site-scripts/index.js ***!
  \****************************************************/
/*! exports provided: SiteScripts */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "NCg5");
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./list */ "Vp/b");
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "b9JD");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteScripts", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["SiteScripts"]; });






Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_2__["SPRest"].prototype, "siteScripts", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_3__["SiteScripts"])(this._baseUrl);
    },
});
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "UuUm":
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/site-groups/types.js ***!
  \***************************************************/
/*! exports provided: _SiteGroups, SiteGroups, _SiteGroup, SiteGroup */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SiteGroups", function() { return _SiteGroups; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SiteGroups", function() { return SiteGroups; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SiteGroup", function() { return _SiteGroup; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SiteGroup", function() { return SiteGroup; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _site_users_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../site-users/types */ "y+KB");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../telemetry */ "h3gW");









var _SiteGroups = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SiteGroups, _super);
    function _SiteGroups() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a group from the collection by id
     *
     * @param id The id of the group to retrieve
     */
    _SiteGroups.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(SiteGroup(this).concat("(" + id + ")"), "sgs.getById");
    };
    /**
     * Adds a new group to the site collection
     *
     * @param properties The group properties object of property names and values to be set for the group
     */
    _SiteGroups.prototype.add = function (properties) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_4__["metadata"])("SP.Group"), properties));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this, "sgs.add"), postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                group: this.getById(data.Id),
                            }];
                }
            });
        });
    };
    /**
     * Gets a group from the collection by name
     *
     * @param groupName The name of the group to retrieve
     */
    _SiteGroups.prototype.getByName = function (groupName) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(SiteGroup(this, "getByName('" + groupName + "')"), "sgs.getByName");
    };
    /**
     * Removes the group with the specified member id from the collection
     *
     * @param id The id of the group to remove
     */
    _SiteGroups.prototype.removeById = function (id) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(SiteGroups, "removeById('" + id + "')"));
    };
    /**
     * Removes the cross-site group with the specified name from the collection
     *
     * @param loginName The name of the group to remove
     */
    _SiteGroups.prototype.removeByLoginName = function (loginName) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(SiteGroups, "removeByLoginName('" + loginName + "')"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("sgs.removeById")
    ], _SiteGroups.prototype, "removeById", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("sgs.removeByLoginName")
    ], _SiteGroups.prototype, "removeByLoginName", null);
    _SiteGroups = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("sitegroups")
    ], _SiteGroups);
    return _SiteGroups;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var SiteGroups = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_SiteGroups);
var _SiteGroup = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SiteGroup, _super);
    function _SiteGroup() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        /**
         * Updates the group with the given property values
         *
         * @param props The group properties object of property names and values to be set for the group
         */
        _this.update = _this._update("SP.Group", function (d, p) {
            var retGroup = _this;
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["hOP"])(p, "Title")) {
                /* tslint:disable-next-line no-string-literal */
                retGroup = _this.getParent(SiteGroup, _this.parentUrl, "getByName('" + p["Title"] + "')");
            }
            return {
                data: d,
                group: retGroup,
            };
        });
        return _this;
    }
    Object.defineProperty(_SiteGroup.prototype, "users", {
        /**
         * Gets the users for this group
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(Object(_site_users_types__WEBPACK_IMPORTED_MODULE_2__["SiteUsers"])(this, "users"), "sg.users");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Set the owner of a group using a user id
     * @param userId the id of the user that will be set as the owner of the current group
     */
    _SiteGroup.prototype.setUserAsOwner = function (userId) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(SiteGroup, "SetUserAsOwner(" + userId + ")"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("sg.setUserAsOwner")
    ], _SiteGroup.prototype, "setUserAsOwner", null);
    return _SiteGroup;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var SiteGroup = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_SiteGroup);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "UxNQ":
/*!*****************************************!*\
  !*** ./node_modules/@pnp/common/net.js ***!
  \*****************************************/
/*! exports provided: mergeHeaders, mergeOptions, getADALResource, FetchClient, BearerTokenFetchClient, SPFxAdalClient */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "mergeHeaders", function() { return mergeHeaders; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "mergeOptions", function() { return mergeOptions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getADALResource", function() { return getADALResource; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "FetchClient", function() { return FetchClient; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "BearerTokenFetchClient", function() { return BearerTokenFetchClient; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SPFxAdalClient", function() { return SPFxAdalClient; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "lzYu");
/* harmony import */ var _util__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./util */ "97IE");
/* harmony import */ var _safe_global__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./safe-global */ "08ZC");



function mergeHeaders(target, source) {
    if (Object(_util__WEBPACK_IMPORTED_MODULE_1__["objectDefinedNotNull"])(source)) {
        var temp = new Request("", { headers: source });
        temp.headers.forEach(function (value, name) {
            target.append(name, value);
        });
    }
}
function mergeOptions(target, source) {
    if (Object(_util__WEBPACK_IMPORTED_MODULE_1__["objectDefinedNotNull"])(source)) {
        var headers = Object(_util__WEBPACK_IMPORTED_MODULE_1__["assign"])(target.headers || {}, source.headers);
        target = Object(_util__WEBPACK_IMPORTED_MODULE_1__["assign"])(target, source);
        target.headers = headers;
    }
}
/**
 * Parses out the root of the request url to use as the resource when getting the token
 *
  * @param url The url to parse
 */
function getADALResource(url) {
    var u = new URL(url);
    return u.protocol + "//" + u.hostname;
}
/**
 * Makes requests using the global/window fetch API
 */
var FetchClient = /** @class */ (function () {
    function FetchClient() {
    }
    FetchClient.prototype.fetch = function (url, options) {
        return _safe_global__WEBPACK_IMPORTED_MODULE_2__["safeGlobal"].fetch(url, options);
    };
    return FetchClient;
}());

/**
 * Makes requests using the fetch API adding the supplied token to the Authorization header
 */
var BearerTokenFetchClient = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(BearerTokenFetchClient, _super);
    function BearerTokenFetchClient(_token) {
        var _this = _super.call(this) || this;
        _this._token = _token;
        return _this;
    }
    Object.defineProperty(BearerTokenFetchClient.prototype, "token", {
        get: function () {
            return this._token || "";
        },
        set: function (token) {
            this._token = token;
        },
        enumerable: true,
        configurable: true
    });
    BearerTokenFetchClient.prototype.fetch = function (url, options) {
        if (options === void 0) { options = {}; }
        var headers = new Headers();
        mergeHeaders(headers, options.headers);
        headers.set("Authorization", "Bearer " + this._token);
        options.headers = headers;
        return _super.prototype.fetch.call(this, url, options);
    };
    return BearerTokenFetchClient;
}(FetchClient));

/**
 * Client wrapping the aadTokenProvider available from SPFx >= 1.6
 */
var SPFxAdalClient = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(SPFxAdalClient, _super);
    /**
     *
     * @param context provide the appropriate SPFx Context object
     */
    function SPFxAdalClient(context) {
        var _this = _super.call(this, null) || this;
        _this.context = context;
        return _this;
    }
    /**
     * Executes a fetch request using the supplied url and options
     *
     * @param url Absolute url of the request
     * @param options Any options
     */
    SPFxAdalClient.prototype.fetch = function (url, options) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var token;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.getToken(getADALResource(url))];
                    case 1:
                        token = _a.sent();
                        this.token = token;
                        return [2 /*return*/, _super.prototype.fetch.call(this, url, options)];
                }
            });
        });
    };
    /**
     * Gets an AAD token for the provided resource using the SPFx AADTokenProvider
     *
     * @param resource Resource for which a token is to be requested (ex: https://graph.microsoft.com)
     */
    SPFxAdalClient.prototype.getToken = function (resource) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var provider;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.context.aadTokenProviderFactory.getTokenProvider()];
                    case 1:
                        provider = _a.sent();
                        return [2 /*return*/, provider.getToken(resource)];
                }
            });
        });
    };
    return SPFxAdalClient;
}(BearerTokenFetchClient));

//# sourceMappingURL=net.js.map

/***/ }),

/***/ "VQtT":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/folder.js ***!
  \************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _folders_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../folders/types */ "4tXH");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "AvPO");



_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"].prototype.shareWith = function (loginNames, role, requireSignin, shareEverything, emailData) {
    if (role === void 0) { role = _types__WEBPACK_IMPORTED_MODULE_2__["SharingRole"].View; }
    if (requireSignin === void 0) { requireSignin = false; }
    if (shareEverything === void 0) { shareEverything = false; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, shareable;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    return [4 /*yield*/, this.getShareable()];
                case 1:
                    shareable = _a.sent();
                    dependency();
                    return [2 /*return*/, shareable.shareWith(loginNames, role, requireSignin, shareEverything, emailData)];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"].prototype.getShareLink = function (kind, expiration) {
    if (expiration === void 0) { expiration = null; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, shareable;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    return [4 /*yield*/, this.getShareable()];
                case 1:
                    shareable = _a.sent();
                    dependency();
                    return [2 /*return*/, shareable.getShareLink(kind, expiration)];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"].prototype.checkSharingPermissions = function (recipients) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, shareable;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    return [4 /*yield*/, this.getShareable()];
                case 1:
                    shareable = _a.sent();
                    dependency();
                    return [2 /*return*/, shareable.checkSharingPermissions(recipients)];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"].prototype.getSharingInformation = function (request, expands) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, shareable;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    return [4 /*yield*/, this.getShareable()];
                case 1:
                    shareable = _a.sent();
                    dependency();
                    return [2 /*return*/, shareable.getSharingInformation(request, expands)];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"].prototype.getObjectSharingSettings = function (useSimplifiedRoles) {
    if (useSimplifiedRoles === void 0) { useSimplifiedRoles = true; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, shareable;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    return [4 /*yield*/, this.getShareable()];
                case 1:
                    shareable = _a.sent();
                    dependency();
                    return [2 /*return*/, shareable.getObjectSharingSettings(useSimplifiedRoles)];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"].prototype.unshare = function () {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, shareable;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    return [4 /*yield*/, this.getShareable()];
                case 1:
                    shareable = _a.sent();
                    dependency();
                    return [2 /*return*/, shareable.unshare()];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"].prototype.deleteSharingLinkByKind = function (kind) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, shareable;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    return [4 /*yield*/, this.getShareable()];
                case 1:
                    shareable = _a.sent();
                    dependency();
                    return [2 /*return*/, shareable.deleteSharingLinkByKind(kind)];
            }
        });
    });
};
_folders_types__WEBPACK_IMPORTED_MODULE_1__["_Folder"].prototype.unshareLink = function (kind, shareId) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var dependency, shareable;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    dependency = this.addBatchDependency();
                    return [4 /*yield*/, this.getShareable()];
                case 1:
                    shareable = _a.sent();
                    dependency();
                    return [2 /*return*/, shareable.unshareLink(kind, shareId)];
            }
        });
    });
};
//# sourceMappingURL=folder.js.map

/***/ }),

/***/ "VnYa":
/*!**********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/list.js ***!
  \**********************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "1mTr");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "userCustomActions", _types__WEBPACK_IMPORTED_MODULE_2__["UserCustomActions"]);
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "Vp/b":
/*!***************************************************!*\
  !*** ./node_modules/@pnp/sp/site-scripts/list.js ***!
  \***************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "b9JD");
/* harmony import */ var _folders_list__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../folders/list */ "vttn");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");







_lists_types__WEBPACK_IMPORTED_MODULE_2__["_List"].prototype.getSiteScript = function () {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var rootFolder, web, absoluteListUrl;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, this.clone(_lists_types__WEBPACK_IMPORTED_MODULE_2__["List"]).rootFolder.get()];
                case 1:
                    rootFolder = _a.sent();
                    return [4 /*yield*/, Object(_webs_types__WEBPACK_IMPORTED_MODULE_5__["Web"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_6__["extractWebUrl"])(this.toUrl())).select("Url").get()];
                case 2:
                    web = _a.sent();
                    absoluteListUrl = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"])(web.Url, "Lists", rootFolder.Name);
                    return [2 /*return*/, Object(_types__WEBPACK_IMPORTED_MODULE_3__["SiteScripts"])(this, "").getSiteScriptFromList(absoluteListUrl)];
            }
        });
    });
};
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "Vpmm":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/logging/logger.js ***!
  \*********************************************/
/*! exports provided: Logger, LogLevel */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Logger", function() { return Logger; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "LogLevel", function() { return LogLevel; });
/**
 * Class used to subscribe ILogListener and log messages throughout an application
 *
 */
var Logger = /** @class */ (function () {
    function Logger() {
    }
    Object.defineProperty(Logger, "activeLogLevel", {
        /**
         * Gets or sets the active log level to apply for log filtering
         */
        get: function () {
            return Logger.instance.activeLogLevel;
        },
        set: function (value) {
            Logger.instance.activeLogLevel = value;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Logger, "instance", {
        get: function () {
            if (Logger._instance === undefined || Logger._instance === null) {
                Logger._instance = new LoggerImpl();
            }
            return Logger._instance;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Adds ILogListener instances to the set of subscribed listeners
     *
     * @param listeners One or more listeners to subscribe to this log
     */
    Logger.subscribe = function () {
        var listeners = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            listeners[_i] = arguments[_i];
        }
        listeners.forEach(function (listener) { return Logger.instance.subscribe(listener); });
    };
    /**
     * Clears the subscribers collection, returning the collection before modification
     */
    Logger.clearSubscribers = function () {
        return Logger.instance.clearSubscribers();
    };
    Object.defineProperty(Logger, "count", {
        /**
         * Gets the current subscriber count
         */
        get: function () {
            return Logger.instance.count;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Writes the supplied string to the subscribed listeners
     *
     * @param message The message to write
     * @param level [Optional] if supplied will be used as the level of the entry (Default: LogLevel.Info)
     */
    Logger.write = function (message, level) {
        if (level === void 0) { level = 1 /* Info */; }
        Logger.instance.log({ level: level, message: message });
    };
    /**
     * Writes the supplied string to the subscribed listeners
     *
     * @param json The json object to stringify and write
     * @param level [Optional] if supplied will be used as the level of the entry (Default: LogLevel.Info)
     */
    Logger.writeJSON = function (json, level) {
        if (level === void 0) { level = 1 /* Info */; }
        this.write(JSON.stringify(json), level);
    };
    /**
     * Logs the supplied entry to the subscribed listeners
     *
     * @param entry The message to log
     */
    Logger.log = function (entry) {
        Logger.instance.log(entry);
    };
    /**
     * Logs an error object to the subscribed listeners
     *
     * @param err The error object
     */
    Logger.error = function (err) {
        Logger.instance.log({ data: err, level: 3 /* Error */, message: err.message });
    };
    return Logger;
}());

var LoggerImpl = /** @class */ (function () {
    function LoggerImpl(activeLogLevel, subscribers) {
        if (activeLogLevel === void 0) { activeLogLevel = 2 /* Warning */; }
        if (subscribers === void 0) { subscribers = []; }
        this.activeLogLevel = activeLogLevel;
        this.subscribers = subscribers;
    }
    LoggerImpl.prototype.subscribe = function (listener) {
        this.subscribers.push(listener);
    };
    LoggerImpl.prototype.clearSubscribers = function () {
        var s = this.subscribers.slice(0);
        this.subscribers.length = 0;
        return s;
    };
    Object.defineProperty(LoggerImpl.prototype, "count", {
        get: function () {
            return this.subscribers.length;
        },
        enumerable: true,
        configurable: true
    });
    LoggerImpl.prototype.write = function (message, level) {
        if (level === void 0) { level = 1 /* Info */; }
        this.log({ level: level, message: message });
    };
    LoggerImpl.prototype.log = function (entry) {
        if (entry !== undefined && this.activeLogLevel <= entry.level) {
            this.subscribers.map(function (subscriber) { return subscriber.log(entry); });
        }
    };
    return LoggerImpl;
}());
/**
 * A set of logging levels
 */
var LogLevel;
(function (LogLevel) {
    LogLevel[LogLevel["Verbose"] = 0] = "Verbose";
    LogLevel[LogLevel["Info"] = 1] = "Info";
    LogLevel[LogLevel["Warning"] = 2] = "Warning";
    LogLevel[LogLevel["Error"] = 3] = "Error";
    LogLevel[LogLevel["Off"] = 99] = "Off";
})(LogLevel || (LogLevel = {}));
//# sourceMappingURL=logger.js.map

/***/ }),

/***/ "W7+7":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/odata/add-prop.js ***!
  \*********************************************/
/*! exports provided: addProp */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "addProp", function() { return addProp; });
/**
 * Adds a property to a target instance
 *
 * @param target The object to whose prototype we will add a property
 * @param name Property name
 * @param factory Factory method used to produce the property value
 * @param path Any additional path required to produce the value
 */
function addProp(target, name, factory, path) {
    Reflect.defineProperty(target.prototype, name, {
        configurable: true,
        enumerable: true,
        get: function () {
            return factory(this, path);
        },
    });
}
//# sourceMappingURL=add-prop.js.map

/***/ }),

/***/ "WqOe":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/hubsites/types.js ***!
  \************************************************/
/*! exports provided: _HubSites, HubSites, _HubSite, HubSite */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_HubSites", function() { return _HubSites; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "HubSites", function() { return HubSites; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_HubSite", function() { return _HubSite; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "HubSite", function() { return HubSite; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _sites_types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sites/types */ "ilAS");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../telemetry */ "h3gW");





var _HubSites = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_HubSites, _super);
    function _HubSites() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a Hub Site from the collection by id
     *
     * @param id The Id of the Hub Site
     */
    _HubSites.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_4__["tag"].configure(HubSite(this, "GetById?hubSiteId='" + id + "'"), "hss.getById");
    };
    _HubSites = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_2__["defaultPath"])("_api/hubsites")
    ], _HubSites);
    return _HubSites;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var HubSites = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_HubSites);
var _HubSite = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_HubSite, _super);
    function _HubSite() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets the ISite instance associated with this hubsite
     */
    _HubSite.prototype.getSite = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var d;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.select("SiteUrl")()];
                    case 1:
                        d = _a.sent();
                        return [2 /*return*/, Object(_sites_types__WEBPACK_IMPORTED_MODULE_3__["Site"])(d.SiteUrl)];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_4__["tag"])("hs.getSite")
    ], _HubSite.prototype, "getSite", null);
    return _HubSite;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var HubSite = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_HubSite);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "XkIg":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/security/list.js ***!
  \***********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "1dmT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./funcs */ "vXSC");





Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "roleAssignments", _types__WEBPACK_IMPORTED_MODULE_2__["RoleAssignments"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "firstUniqueAncestorSecurableObject", _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"]);
_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"].prototype.getUserEffectivePermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["getUserEffectivePermissions"];
_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"].prototype.getCurrentUserEffectivePermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["getCurrentUserEffectivePermissions"];
_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"].prototype.breakRoleInheritance = _funcs__WEBPACK_IMPORTED_MODULE_4__["breakRoleInheritance"];
_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"].prototype.resetRoleInheritance = _funcs__WEBPACK_IMPORTED_MODULE_4__["resetRoleInheritance"];
_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"].prototype.userHasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["userHasPermissions"];
_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"].prototype.currentUserHasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["currentUserHasPermissions"];
_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"].prototype.hasPermissions = _funcs__WEBPACK_IMPORTED_MODULE_4__["hasPermissions"];
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "YFzv":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/file-names.js ***!
  \**************************************************/
/*! exports provided: containsInvalidFileFolderChars, stripInvalidFileFolderChars */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "containsInvalidFileFolderChars", function() { return containsInvalidFileFolderChars; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "stripInvalidFileFolderChars", function() { return stripInvalidFileFolderChars; });
var InvalidFileFolderNameCharsOnlineRegex = /["*:<>?/\\|\x00-\x1f\x7f-\x9f]/g;
var InvalidFileFolderNameCharsOnPremiseRegex = /["#%*:<>?/\\|\x00-\x1f\x7f-\x9f]/g;
/**
 * Checks if file or folder name contains invalid characters
 *
 * @param input File or folder name to check
 * @param onPremise Set to true for SharePoint On-Premise
 * @returns True if contains invalid chars, false otherwise
 */
function containsInvalidFileFolderChars(input, onPremise) {
    if (onPremise === void 0) { onPremise = false; }
    if (onPremise) {
        return InvalidFileFolderNameCharsOnPremiseRegex.test(input);
    }
    else {
        return InvalidFileFolderNameCharsOnlineRegex.test(input);
    }
}
/**
 * Removes invalid characters from file or folder name
 *
 * @param input File or folder name
 * @param replacer Value that will replace invalid characters
 * @param onPremise Set to true for SharePoint On-Premise
 * @returns File or folder name with replaced invalid characters
 */
function stripInvalidFileFolderChars(input, replacer, onPremise) {
    if (replacer === void 0) { replacer = ""; }
    if (onPremise === void 0) { onPremise = false; }
    if (onPremise) {
        return input.replace(InvalidFileFolderNameCharsOnPremiseRegex, replacer);
    }
    else {
        return input.replace(InvalidFileFolderNameCharsOnlineRegex, replacer);
    }
}
//# sourceMappingURL=file-names.js.map

/***/ }),

/***/ "YpSs":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/search/types.js ***!
  \**********************************************/
/*! exports provided: SortDirection, ReorderingRuleMatchType, QueryPropertyValueType, SearchBuiltInSourceId */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SortDirection", function() { return SortDirection; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ReorderingRuleMatchType", function() { return ReorderingRuleMatchType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "QueryPropertyValueType", function() { return QueryPropertyValueType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SearchBuiltInSourceId", function() { return SearchBuiltInSourceId; });
/**
 * defines the SortDirection enum
 */
var SortDirection;
(function (SortDirection) {
    SortDirection[SortDirection["Ascending"] = 0] = "Ascending";
    SortDirection[SortDirection["Descending"] = 1] = "Descending";
    SortDirection[SortDirection["FQLFormula"] = 2] = "FQLFormula";
})(SortDirection || (SortDirection = {}));
/**
 * defines the ReorderingRuleMatchType  enum
 */
var ReorderingRuleMatchType;
(function (ReorderingRuleMatchType) {
    ReorderingRuleMatchType[ReorderingRuleMatchType["ResultContainsKeyword"] = 0] = "ResultContainsKeyword";
    ReorderingRuleMatchType[ReorderingRuleMatchType["TitleContainsKeyword"] = 1] = "TitleContainsKeyword";
    ReorderingRuleMatchType[ReorderingRuleMatchType["TitleMatchesKeyword"] = 2] = "TitleMatchesKeyword";
    ReorderingRuleMatchType[ReorderingRuleMatchType["UrlStartsWith"] = 3] = "UrlStartsWith";
    ReorderingRuleMatchType[ReorderingRuleMatchType["UrlExactlyMatches"] = 4] = "UrlExactlyMatches";
    ReorderingRuleMatchType[ReorderingRuleMatchType["ContentTypeIs"] = 5] = "ContentTypeIs";
    ReorderingRuleMatchType[ReorderingRuleMatchType["FileExtensionMatches"] = 6] = "FileExtensionMatches";
    ReorderingRuleMatchType[ReorderingRuleMatchType["ResultHasTag"] = 7] = "ResultHasTag";
    ReorderingRuleMatchType[ReorderingRuleMatchType["ManualCondition"] = 8] = "ManualCondition";
})(ReorderingRuleMatchType || (ReorderingRuleMatchType = {}));
/**
 * Specifies the type value for the property
 */
var QueryPropertyValueType;
(function (QueryPropertyValueType) {
    QueryPropertyValueType[QueryPropertyValueType["None"] = 0] = "None";
    QueryPropertyValueType[QueryPropertyValueType["StringType"] = 1] = "StringType";
    QueryPropertyValueType[QueryPropertyValueType["Int32Type"] = 2] = "Int32Type";
    QueryPropertyValueType[QueryPropertyValueType["BooleanType"] = 3] = "BooleanType";
    QueryPropertyValueType[QueryPropertyValueType["StringArrayType"] = 4] = "StringArrayType";
    QueryPropertyValueType[QueryPropertyValueType["UnSupportedType"] = 5] = "UnSupportedType";
})(QueryPropertyValueType || (QueryPropertyValueType = {}));
var SearchBuiltInSourceId = /** @class */ (function () {
    function SearchBuiltInSourceId() {
    }
    SearchBuiltInSourceId.Documents = "e7ec8cee-ded8-43c9-beb5-436b54b31e84";
    SearchBuiltInSourceId.ItemsMatchingContentType = "5dc9f503-801e-4ced-8a2c-5d1237132419";
    SearchBuiltInSourceId.ItemsMatchingTag = "e1327b9c-2b8c-4b23-99c9-3730cb29c3f7";
    SearchBuiltInSourceId.ItemsRelatedToCurrentUser = "48fec42e-4a92-48ce-8363-c2703a40e67d";
    SearchBuiltInSourceId.ItemsWithSameKeywordAsThisItem = "5c069288-1d17-454a-8ac6-9c642a065f48";
    SearchBuiltInSourceId.LocalPeopleResults = "b09a7990-05ea-4af9-81ef-edfab16c4e31";
    SearchBuiltInSourceId.LocalReportsAndDataResults = "203fba36-2763-4060-9931-911ac8c0583b";
    SearchBuiltInSourceId.LocalSharePointResults = "8413cd39-2156-4e00-b54d-11efd9abdb89";
    SearchBuiltInSourceId.LocalVideoResults = "78b793ce-7956-4669-aa3b-451fc5defebf";
    SearchBuiltInSourceId.Pages = "5e34578e-4d08-4edc-8bf3-002acf3cdbcc";
    SearchBuiltInSourceId.Pictures = "38403c8c-3975-41a8-826e-717f2d41568a";
    SearchBuiltInSourceId.Popular = "97c71db1-58ce-4891-8b64-585bc2326c12";
    SearchBuiltInSourceId.RecentlyChangedItems = "ba63bbae-fa9c-42c0-b027-9a878f16557c";
    SearchBuiltInSourceId.RecommendedItems = "ec675252-14fa-4fbe-84dd-8d098ed74181";
    SearchBuiltInSourceId.Wiki = "9479bf85-e257-4318-b5a8-81a180f5faa1";
    return SearchBuiltInSourceId;
}());

//# sourceMappingURL=types.js.map

/***/ }),

/***/ "ZdGM":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/navigation/index.js ***!
  \**************************************************/
/*! exports provided: Navigation, NavigationNode, NavigationNodes, NavigationService */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "BK4w");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./web */ "pTlN");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Navigation", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Navigation"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "NavigationNode", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["NavigationNode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "NavigationNodes", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["NavigationNodes"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "NavigationService", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["NavigationService"]; });





Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"].prototype, "navigation", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_1__["NavigationService"])().configure(this._options);
    },
});
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "Zl34":
/*!********************************************!*\
  !*** ./node_modules/@pnp/odata/caching.js ***!
  \********************************************/
/*! exports provided: CachingOptions, CachingParserWrapper */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CachingOptions", function() { return CachingOptions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CachingParserWrapper", function() { return CachingParserWrapper; });
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/common */ "zqEK");

var CachingOptions = /** @class */ (function () {
    function CachingOptions(key, storeName, expiration) {
        if (storeName === void 0) { storeName = _pnp_common__WEBPACK_IMPORTED_MODULE_0__["RuntimeConfig"].defaultCachingStore; }
        if (expiration === void 0) { expiration = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["dateAdd"])(new Date(), "second", _pnp_common__WEBPACK_IMPORTED_MODULE_0__["RuntimeConfig"].defaultCachingTimeoutSeconds); }
        this.key = key;
        this.storeName = storeName;
        this.expiration = expiration;
    }
    Object.defineProperty(CachingOptions.prototype, "store", {
        get: function () {
            if (this.storeName === "local") {
                return CachingOptions.storage.local;
            }
            else {
                return CachingOptions.storage.session;
            }
        },
        enumerable: true,
        configurable: true
    });
    CachingOptions.storage = new _pnp_common__WEBPACK_IMPORTED_MODULE_0__["PnPClientStorage"]();
    return CachingOptions;
}());

var CachingParserWrapper = /** @class */ (function () {
    function CachingParserWrapper(parser, cacheOptions) {
        this.parser = parser;
        this.cacheOptions = cacheOptions;
    }
    CachingParserWrapper.prototype.parse = function (response) {
        var _this = this;
        return this.parser.parse(response).then(function (r) { return _this.cacheData(r); });
    };
    CachingParserWrapper.prototype.cacheData = function (data) {
        if (this.cacheOptions.store !== null) {
            this.cacheOptions.store.put(this.cacheOptions.key, data, this.cacheOptions.expiration);
        }
        return data;
    };
    return CachingParserWrapper;
}());

//# sourceMappingURL=caching.js.map

/***/ }),

/***/ "Zlzk":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/features/site.js ***!
  \***********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sites_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sites/types */ "ilAS");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "HHhI");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_sites_types__WEBPACK_IMPORTED_MODULE_1__["_Site"], "features", _types__WEBPACK_IMPORTED_MODULE_2__["Features"]);
//# sourceMappingURL=site.js.map

/***/ }),

/***/ "aH1g":
/*!************************************************!*\
  !*** ./node_modules/@pnp/logging/listeners.js ***!
  \************************************************/
/*! exports provided: ConsoleListener, FunctionListener */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ConsoleListener", function() { return ConsoleListener; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "FunctionListener", function() { return FunctionListener; });
/**
 * Implementation of LogListener which logs to the console
 *
 */
var ConsoleListener = /** @class */ (function () {
    function ConsoleListener() {
    }
    /**
     * Any associated data that a given logging listener may choose to log or ignore
     *
     * @param entry The information to be logged
     */
    ConsoleListener.prototype.log = function (entry) {
        var msg = this.format(entry);
        switch (entry.level) {
            case 0 /* Verbose */:
            case 1 /* Info */:
                console.log(msg);
                break;
            case 2 /* Warning */:
                console.warn(msg);
                break;
            case 3 /* Error */:
                console.error(msg);
                break;
        }
    };
    /**
     * Formats the message
     *
     * @param entry The information to format into a string
     */
    ConsoleListener.prototype.format = function (entry) {
        var msg = [];
        msg.push("Message: " + entry.message);
        if (entry.data !== undefined) {
            try {
                msg.push(" Data: " + JSON.stringify(entry.data));
            }
            catch (e) {
                msg.push(" Data: Error in stringify of supplied data " + e);
            }
        }
        return msg.join("");
    };
    return ConsoleListener;
}());

/**
 * Implementation of LogListener which logs to the supplied function
 *
 */
var FunctionListener = /** @class */ (function () {
    /**
     * Creates a new instance of the FunctionListener class
     *
     * @constructor
     * @param  method The method to which any logging data will be passed
     */
    function FunctionListener(method) {
        this.method = method;
    }
    /**
     * Any associated data that a given logging listener may choose to log or ignore
     *
     * @param entry The information to be logged
     */
    FunctionListener.prototype.log = function (entry) {
        this.method(entry);
    };
    return FunctionListener;
}());

//# sourceMappingURL=listeners.js.map

/***/ }),

/***/ "ah5w":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/security/index.js ***!
  \************************************************/
/*! exports provided: RoleAssignment, RoleAssignments, RoleDefinition, RoleDefinitions, PermissionKind */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _item__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./item */ "Ka72");
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./list */ "XkIg");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./web */ "4G66");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "1dmT");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleAssignment", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["RoleAssignment"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleAssignments", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["RoleAssignments"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleDefinition", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["RoleDefinition"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleDefinitions", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["RoleDefinitions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PermissionKind", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["PermissionKind"]; });





//# sourceMappingURL=index.js.map

/***/ }),

/***/ "amhO":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/appcatalog/web.js ***!
  \************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "ArVI");


_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getAppCatalog = function (url) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["AppCatalog"])(url || this);
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "auja":
/*!****************************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/content-type.js ***!
  \****************************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _content_types_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../content-types/types */ "S3gC");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./funcs */ "T2j4");


_content_types_types__WEBPACK_IMPORTED_MODULE_0__["_ContentType"].prototype.titleResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_1__["getValueForUICultureBinder"])("titleResource");
_content_types_types__WEBPACK_IMPORTED_MODULE_0__["_ContentType"].prototype.descriptionResource = Object(_funcs__WEBPACK_IMPORTED_MODULE_1__["getValueForUICultureBinder"])("descriptionResource");
//# sourceMappingURL=content-type.js.map

/***/ }),

/***/ "b/a5":
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/list.js ***!
  \****************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "S3gC");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "contentTypes", _types__WEBPACK_IMPORTED_MODULE_2__["ContentTypes"]);
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "b9JD":
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/site-scripts/types.js ***!
  \****************************************************/
/*! exports provided: _SiteScripts, SiteScripts, SiteScriptActionOutcome */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SiteScripts", function() { return _SiteScripts; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SiteScripts", function() { return SiteScripts; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SiteScriptActionOutcome", function() { return SiteScriptActionOutcome; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");







var _SiteScripts = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SiteScripts, _super);
    function _SiteScripts(baseUrl, methodName) {
        if (methodName === void 0) { methodName = ""; }
        var _this = this;
        var url = typeof baseUrl === "string" ? baseUrl : baseUrl.toUrl();
        _this = _super.call(this, Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_4__["extractWebUrl"])(url), "_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + methodName) || this;
        return _this;
    }
    _SiteScripts.prototype.execute = function (props) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_2__["spPost"])(this, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])(props));
    };
    /**
     * Gets a list of information on all existing site scripts.
     */
    _SiteScripts.prototype.getSiteScripts = function () {
        return this.clone(SiteScriptsCloneFactory, "GetSiteScripts", true).execute({});
    };
    /**
     * Creates a new site script.
     *
     * @param title The display name of the site script.
     * @param content JSON value that describes the script. For more information, see JSON reference.
     */
    _SiteScripts.prototype.createSiteScript = function (title, description, content) {
        return this.clone(SiteScriptsCloneFactory, "CreateSiteScript(Title=@title,Description=@desc)?@title='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_6__["escapeQueryStrValue"])(title) + "'&@desc='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_6__["escapeQueryStrValue"])(description) + "'")
            .execute(content);
    };
    /**
     * Gets information about a specific site script. It also returns the JSON of the script.
     *
     * @param id The ID of the site script to get information about.
     */
    _SiteScripts.prototype.getSiteScriptMetadata = function (id) {
        return this.clone(SiteScriptsCloneFactory, "GetSiteScriptMetadata").execute({ id: id });
    };
    /**
     * Deletes a site script.
     *
     * @param id The ID of the site script to delete.
     */
    _SiteScripts.prototype.deleteSiteScript = function (id) {
        return this.clone(SiteScriptsCloneFactory, "DeleteSiteScript").execute({ id: id });
    };
    /**
     * Updates a site script with new values. In the REST call, all parameters are optional except the site script Id.
     *
     * @param siteScriptUpdateInfo Object that contains the information to update a site script.
     *                             Make sure you stringify the content object or pass it in the second 'content' parameter
     * @param content (Optional) A new JSON script defining the script actions. For more information, see Site design JSON schema.
     */
    _SiteScripts.prototype.updateSiteScript = function (updateInfo, content) {
        if (content) {
            updateInfo.Content = JSON.stringify(content);
        }
        return this.clone(SiteScriptsCloneFactory, "UpdateSiteScript").execute({ updateInfo: updateInfo });
    };
    /**
     * Gets the site script syntax (JSON) for a specific list
     * @param listUrl The absolute url of the list to retrieve site script
     */
    _SiteScripts.prototype.getSiteScriptFromList = function (listUrl) {
        return this.clone(SiteScriptsCloneFactory, "GetSiteScriptFromList").execute({ listUrl: listUrl });
    };
    /**
     * Gets the site script syntax (JSON) for a specific web
     * @param webUrl The absolute url of the web to retrieve site script
     * @param extractInfo configuration object to specify what to extract
     */
    _SiteScripts.prototype.getSiteScriptFromWeb = function (webUrl, info) {
        return this.clone(SiteScriptsCloneFactory, "getSiteScriptFromWeb").execute({ webUrl: webUrl, info: info });
    };
    /**
     * Executes the indicated site design action on the indicated web.
     *
     * @param webUrl The absolute url of the web to retrieve site script
     * @param extractInfo configuration object to specify what to extract
     */
    _SiteScripts.prototype.executeSiteScriptAction = function (actionDefinition) {
        return this.clone(SiteScriptsCloneFactory, "executeSiteScriptAction").execute({ actionDefinition: actionDefinition });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ss.getSiteScripts")
    ], _SiteScripts.prototype, "getSiteScripts", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ss.createSiteScript")
    ], _SiteScripts.prototype, "createSiteScript", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ss.getSiteScriptMetadata")
    ], _SiteScripts.prototype, "getSiteScriptMetadata", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ss.deleteSiteScript")
    ], _SiteScripts.prototype, "deleteSiteScript", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ss.updateSiteScript")
    ], _SiteScripts.prototype, "updateSiteScript", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ss.getSiteScriptFromList")
    ], _SiteScripts.prototype, "getSiteScriptFromList", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ss.getSiteScriptFromWeb")
    ], _SiteScripts.prototype, "getSiteScriptFromWeb", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_5__["tag"])("ss.executeSiteScriptAction")
    ], _SiteScripts.prototype, "executeSiteScriptAction", null);
    return _SiteScripts;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryable"]));

var SiteScripts = function (baseUrl, methodName) { return new _SiteScripts(baseUrl, methodName); };
var SiteScriptsCloneFactory = function (baseUrl, methodName) {
    if (methodName === void 0) { methodName = ""; }
    return SiteScripts(baseUrl, methodName);
};
var SiteScriptActionOutcome;
(function (SiteScriptActionOutcome) {
    /**
     * The stage was deemed to have completed successfully.
     */
    SiteScriptActionOutcome[SiteScriptActionOutcome["Success"] = 0] = "Success";
    /**
     * The stage was deemed to have failed to complete successfully (non-blocking, rest of recipe
     * execution should still be able to proceed).
     */
    SiteScriptActionOutcome[SiteScriptActionOutcome["Failure"] = 1] = "Failure";
    /**
     * No action was taken for this stage / this stage was skipped.
     */
    SiteScriptActionOutcome[SiteScriptActionOutcome["NoOp"] = 2] = "NoOp";
    /**
     * There was an exception but the operation succeeded. This is analagous to the operation completing
     * in a "yellow" state.
     */
    SiteScriptActionOutcome[SiteScriptActionOutcome["SucceededWithException"] = 3] = "SucceededWithException";
})(SiteScriptActionOutcome || (SiteScriptActionOutcome = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "cDcd":
/*!************************!*\
  !*** external "react" ***!
  \************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_cDcd__;

/***/ }),

/***/ "dGPD":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/content-types/index.js ***!
  \*****************************************************/
/*! exports provided: ContentType, ContentTypes, FieldLink, FieldLinks */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "M9RU");
/* harmony import */ var _item__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./item */ "4Egx");
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./list */ "b/a5");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./types */ "S3gC");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ContentType", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["ContentType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ContentTypes", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["ContentTypes"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FieldLink", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["FieldLink"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FieldLinks", function() { return _types__WEBPACK_IMPORTED_MODULE_3__["FieldLinks"]; });





//# sourceMappingURL=index.js.map

/***/ }),

/***/ "dKLh":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/odata/queryable.js ***!
  \**********************************************/
/*! exports provided: cloneQueryableData, Queryable */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "cloneQueryableData", function() { return cloneQueryableData; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Queryable", function() { return Queryable; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "t3D1");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _parsers__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./parsers */ "xbXq");



function cloneQueryableData(source) {
    var s = JSON.stringify(source, function (key, value) {
        switch (key) {
            case "query":
                return JSON.stringify(Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])(value));
            case "batch":
                return "-";
            case "batchDependency":
                return "-";
            case "cachingOptions":
                return "-";
            case "clientFactory":
                return "-";
            case "parser":
                return "-";
            default:
                return value;
        }
    }, 0);
    var parsed = JSON.parse(s, function (key, value) {
        switch (key) {
            case "query":
                return new Map(JSON.parse(value));
            case "batch":
                return source.batch;
            case "batchDependency":
                return source.batchDependency;
            case "cachingOptions":
                return source.cachingOptions;
            case "clientFactory":
                return source.clientFactory;
            case "parser":
                return source.parser;
            default:
                return value;
        }
    });
    // this handles bodies that cannot be JSON encoded (Blob, etc)
    if (source.options && source.options.body) {
        parsed.options.body = source.options.body;
    }
    return parsed;
}
var Queryable = /** @class */ (function () {
    function Queryable(dataSeed) {
        if (dataSeed === void 0) { dataSeed = {}; }
        this._data = Object.assign({}, {
            cloneParentWasCaching: false,
            options: {},
            parentUrl: "",
            parser: new _parsers__WEBPACK_IMPORTED_MODULE_2__["ODataParser"](),
            query: new Map(),
            url: "",
            useCaching: false,
        }, cloneQueryableData(dataSeed));
    }
    Object.defineProperty(Queryable.prototype, "data", {
        get: function () {
            return this._data;
        },
        set: function (value) {
            this._data = Object.assign({}, cloneQueryableData(this.data), cloneQueryableData(value));
        },
        enumerable: true,
        configurable: true
    });
    /**
    * Gets the currentl url
    *
    */
    Queryable.prototype.toUrl = function () {
        return this.data.url;
    };
    /**
     * Directly concatenates the supplied string to the current url, not normalizing "/" chars
     *
     * @param pathPart The string to concatenate to the url
     */
    Queryable.prototype.concat = function (pathPart) {
        this.data.url += pathPart;
        return this;
    };
    Object.defineProperty(Queryable.prototype, "query", {
        /**
         * Provides access to the query builder for this url
         *
         */
        get: function () {
            return this.data.query;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Sets custom options for current object and all derived objects accessible via chaining
     *
     * @param options custom options
     */
    Queryable.prototype.configure = function (options) {
        Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["mergeOptions"])(this.data.options, options);
        return this;
    };
    /**
     * Configures this instance from the configure options of the supplied instance
     *
     * @param o Instance from which options should be taken
     */
    Queryable.prototype.configureFrom = function (o) {
        Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["mergeOptions"])(this.data.options, o.data.options);
        return this;
    };
    /**
     * Enables caching for this request
     *
     * @param options Defines the options used when caching this request
     */
    Queryable.prototype.usingCaching = function (options) {
        if (!_pnp_common__WEBPACK_IMPORTED_MODULE_1__["RuntimeConfig"].globalCacheDisable) {
            this.data.useCaching = true;
            if (options !== undefined) {
                this.data.cachingOptions = options;
            }
        }
        return this;
    };
    Queryable.prototype.usingParser = function (parser) {
        this.data.parser = parser;
        return this;
    };
    /**
     * Allows you to set a request specific processing pipeline
     *
     * @param pipeline The set of methods, in order, to execute a given request
     */
    Queryable.prototype.withPipeline = function (pipeline) {
        this.data.pipes = pipeline.slice(0);
        return this;
    };
    /**
     * Appends the given string and normalizes "/" chars
     *
     * @param pathPart The string to append
     */
    Queryable.prototype.append = function (pathPart) {
        this.data.url = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["combine"])(this.data.url, pathPart);
    };
    /**
     * Adds this query to the supplied batch
     *
     * @example
     * ```
     *
     * let b = pnp.sp.createBatch();
     * pnp.sp.web.inBatch(b).get().then(...);
     * b.execute().then(...)
     * ```
     */
    Queryable.prototype.inBatch = function (batch) {
        if (this.hasBatch) {
            throw Error("This query is already part of a batch.");
        }
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["objectDefinedNotNull"])(batch)) {
            batch.track(this);
        }
        return this;
    };
    /**
     * Blocks a batch call from occuring, MUST be cleared by calling the returned function
    */
    Queryable.prototype.addBatchDependency = function () {
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["objectDefinedNotNull"])(this.data.batch)) {
            return this.data.batch.addDependency();
        }
        return function () { return null; };
    };
    Object.defineProperty(Queryable.prototype, "hasBatch", {
        /**
         * Indicates if the current query has a batch associated
         *
         */
        get: function () {
            return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["objectDefinedNotNull"])(this.data.batch);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Queryable.prototype, "batch", {
        /**
         * The batch currently associated with this query or null
         *
         */
        get: function () {
            return this.hasBatch ? this.data.batch : null;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Queryable.prototype, "parentUrl", {
        /**
         * Gets the parent url used when creating this instance
         *
         */
        get: function () {
            return this.data.parentUrl;
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Clones this instance's data to target
     *
     * @param target Instance to which data is written
     * @param settings [Optional] Settings controlling how clone is applied
     */
    Queryable.prototype.cloneTo = function (target, settings) {
        if (settings === void 0) { settings = { includeBatch: true }; }
        target.data = Object.assign({}, cloneQueryableData(this.data), {
            batch: null,
            cloneParentCacheOptions: null,
            cloneParentWasCaching: false,
        }, cloneQueryableData(target.data));
        target.configureFrom(this);
        if (settings.includeBatch) {
            target.inBatch(this.batch);
        }
        if (this.data.useCaching) {
            target.data.cloneParentWasCaching = true;
            target.data.cloneParentCacheOptions = this.data.cachingOptions;
        }
        return target;
    };
    return Queryable;
}());

//# sourceMappingURL=queryable.js.map

/***/ }),

/***/ "dRNj":
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/subscriptions/list.js ***!
  \****************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "N3o1");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "subscriptions", _types__WEBPACK_IMPORTED_MODULE_2__["Subscriptions"]);
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "dVsc":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/webs/types.js ***!
  \********************************************/
/*! exports provided: _Webs, Webs, _Web, Web */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Webs", function() { return _Webs; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Webs", function() { return Webs; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Web", function() { return _Web; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Web", function() { return Web; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _batch__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../batch */ "8xxP");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _sites__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../sites */ "1jZ6");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../telemetry */ "h3gW");












var _Webs = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Webs, _super);
    function _Webs() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Adds a new web to the collection
     *
     * @param title The new web's title
     * @param url The new web's relative url
     * @param description The new web's description
     * @param template The new web's template internal name (default = STS)
     * @param language The locale id that specifies the new web's language (default = 1033 [English, US])
     * @param inheritPermissions When true, permissions will be inherited from the new web's parent (default = true)
     */
    _Webs.prototype.add = function (title, url, description, template, language, inheritPermissions) {
        if (description === void 0) { description = ""; }
        if (template === void 0) { template = "STS"; }
        if (language === void 0) { language = 1033; }
        if (inheritPermissions === void 0) { inheritPermissions = true; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({
                            "parameters": Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_7__["metadata"])("SP.WebCreationInformation"), {
                                Description: description,
                                Language: language,
                                Title: title,
                                Url: url,
                                UseSamePermissionsAsParentSite: inheritPermissions,
                                WebTemplate: template,
                            }),
                        });
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spPost"])(this.clone(Webs, "add"), postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                web: Web(Object(_odata__WEBPACK_IMPORTED_MODULE_5__["odataUrlFrom"])(data).replace(/_api\/web\/?/i, "")),
                            }];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("ws.add")
    ], _Webs.prototype, "add", null);
    _Webs = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_4__["defaultPath"])("webs")
    ], _Webs);
    return _Webs;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableCollection"]));

var Webs = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_Webs);
/**
 * Describes a web
 *
 */
var _Web = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Web, _super);
    function _Web() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["deleteable"])("w");
        return _this;
    }
    Object.defineProperty(_Web.prototype, "webs", {
        /**
         * Gets this web's subwebs
         *
         */
        get: function () {
            return Webs(this);
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Gets this web's parent web and data
     *
     */
    _Web.prototype.getParentWeb = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var ParentWeb;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spGet"])(this.select("ParentWeb/Id").expand("ParentWeb"))];
                    case 1:
                        ParentWeb = (_a.sent()).ParentWeb;
                        return [2 /*return*/, ParentWeb ? Object(_sites__WEBPACK_IMPORTED_MODULE_8__["Site"])(this.parentUrl).openWebById(ParentWeb.Id) : null];
                }
            });
        });
    };
    /**
    * Returns a collection of objects that contain metadata about subsites of the current site in which the current user is a member.
    *
    * @param nWebTemplateFilter Specifies the site definition (default = -1)
    * @param nConfigurationFilter A 16-bit integer that specifies the identifier of a configuration (default = -1)
    */
    _Web.prototype.getSubwebsFilteredForCurrentUser = function (nWebTemplateFilter, nConfigurationFilter) {
        if (nWebTemplateFilter === void 0) { nWebTemplateFilter = -1; }
        if (nConfigurationFilter === void 0) { nConfigurationFilter = -1; }
        var o = this.clone(Webs, "getSubwebsFilteredForCurrentUser(nWebTemplateFilter=" + nWebTemplateFilter + ",nConfigurationFilter=" + nConfigurationFilter + ")");
        return _telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"].configure(o, "w.getSubwebsFilteredForCurrentUser");
    };
    Object.defineProperty(_Web.prototype, "allProperties", {
        /**
         * Allows access to the web's all properties collection
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"].configure(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableInstance"], "allproperties"), "w.allprops");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_Web.prototype, "webinfos", {
        /**
         * Gets a collection of WebInfos for this web's subwebs
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableCollection"])(this, "webinfos"), "w.webinfos");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Creates a new batch for requests within the context of this web
     *
     */
    _Web.prototype.createBatch = function () {
        return new _batch__WEBPACK_IMPORTED_MODULE_6__["SPBatch"](this.parentUrl);
    };
    /**
     * Updates this web instance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the web
     */
    _Web.prototype.update = function (properties) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_7__["metadata"])("SP.Web"), properties), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "X-HTTP-Method": "MERGE" }));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spPost"])(this, postBody)];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, { data: data, web: this }];
                }
            });
        });
    };
    /**
     * Applies the theme specified by the contents of each of the files specified in the arguments to the site
     *
     * @param colorPaletteUrl The server-relative URL of the color palette file
     * @param fontSchemeUrl The server-relative URL of the font scheme
     * @param backgroundImageUrl The server-relative URL of the background image
     * @param shareGenerated When true, the generated theme files are stored in the root site. When false, they are stored in this web
     */
    _Web.prototype.applyTheme = function (colorPaletteUrl, fontSchemeUrl, backgroundImageUrl, shareGenerated) {
        var postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({
            backgroundImageUrl: backgroundImageUrl,
            colorPaletteUrl: colorPaletteUrl,
            fontSchemeUrl: fontSchemeUrl,
            shareGenerated: shareGenerated,
        });
        return Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spPost"])(this.clone(Web, "applytheme"), postBody);
    };
    /**
     * Applies the specified site definition or site template to the Web site that has no template applied to it
     *
     * @param template Name of the site definition or the name of the site template
     */
    _Web.prototype.applyWebTemplate = function (template) {
        var q = this.clone(Web, "applywebtemplate");
        q.concat("(webTemplate='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_10__["escapeQueryStrValue"])(template) + "')");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spPost"])(q);
    };
    /**
     * Returns a collection of site templates available for the site
     *
     * @param language The locale id of the site templates to retrieve (default = 1033 [English, US])
     * @param includeCrossLanguage When true, includes language-neutral site templates; otherwise false (default = true)
     */
    _Web.prototype.availableWebTemplates = function (language, includeCrossLanugage) {
        if (language === void 0) { language = 1033; }
        if (includeCrossLanugage === void 0) { includeCrossLanugage = true; }
        var path = "getavailablewebtemplates(lcid=" + language + ", doincludecrosslanguage=" + includeCrossLanugage + ")";
        return _telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableCollection"])(this, path), "w.availableWebTemplates");
    };
    /**
     * Returns the collection of changes from the change log that have occurred within the list, based on the specified query
     *
     * @param query The change query
     */
    _Web.prototype.getChanges = function (query) {
        var postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({ "query": Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_7__["metadata"])("SP.ChangeQuery"), query) });
        return Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spPost"])(this.clone(Web, "getchanges"), postBody);
    };
    /**
     * Returns the name of the image file for the icon that is used to represent the specified file
     *
     * @param filename The file name. If this parameter is empty, the server returns an empty string
     * @param size The size of the icon: 16x16 pixels = 0, 32x32 pixels = 1 (default = 0)
     * @param progId The ProgID of the application that was used to create the file, in the form OLEServerName.ObjectName
     */
    _Web.prototype.mapToIcon = function (filename, size, progId) {
        if (size === void 0) { size = 0; }
        if (progId === void 0) { progId = ""; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spGet"])(this.clone(Web, "maptoicon(filename='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_10__["escapeQueryStrValue"])(filename) + "', progid='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_10__["escapeQueryStrValue"])(progId) + "', size=" + size + ")"));
    };
    /**
     * Returns the tenant property corresponding to the specified key in the app catalog site
     *
     * @param key Id of storage entity to be set
     */
    _Web.prototype.getStorageEntity = function (key) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spGet"])(this.clone(Web, "getStorageEntity('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_10__["escapeQueryStrValue"])(key) + "')"));
    };
    /**
     * This will set the storage entity identified by the given key (MUST be called in the context of the app catalog)
     *
     * @param key Id of storage entity to be set
     * @param value Value of storage entity to be set
     * @param description Description of storage entity to be set
     * @param comments Comments of storage entity to be set
     */
    _Web.prototype.setStorageEntity = function (key, value, description, comments) {
        if (description === void 0) { description = ""; }
        if (comments === void 0) { comments = ""; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spPost"])(this.clone(Web, "setStorageEntity"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({
            comments: comments,
            description: description,
            key: key,
            value: value,
        }));
    };
    /**
     * This will remove the storage entity identified by the given key
     *
     * @param key Id of storage entity to be removed
     */
    _Web.prototype.removeStorageEntity = function (key) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_9__["spPost"])(this.clone(Web, "removeStorageEntity('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_10__["escapeQueryStrValue"])(key) + "')"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.getParentWeb")
    ], _Web.prototype, "getParentWeb", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.update")
    ], _Web.prototype, "update", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.applyTheme")
    ], _Web.prototype, "applyTheme", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.applyWebTemplate")
    ], _Web.prototype, "applyWebTemplate", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.getChanges")
    ], _Web.prototype, "getChanges", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.mapToIcon")
    ], _Web.prototype, "mapToIcon", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.getStorageEntity")
    ], _Web.prototype, "getStorageEntity", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.setStorageEntity")
    ], _Web.prototype, "setStorageEntity", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_11__["tag"])("w.removeStorageEntity")
    ], _Web.prototype, "removeStorageEntity", null);
    _Web = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_4__["defaultPath"])("_api/web")
    ], _Web);
    return _Web;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableInstance"]));

var Web = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_Web);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "dYBM":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/views/index.js ***!
  \*********************************************/
/*! exports provided: View, ViewFields, Views, ViewScope */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list */ "iV2g");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "32VY");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "View", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["View"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ViewFields", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ViewFields"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Views", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Views"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ViewScope", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ViewScope"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "eDdy":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/site-designs/web.js ***!
  \**************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "w4d6");


_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getSiteDesignRuns = function (siteDesignId) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["SiteDesigns"])(this, "").getSiteDesignRun(undefined, siteDesignId);
};
_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.addSiteDesignTask = function (siteDesignId) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["SiteDesigns"])(this, "").addSiteDesignTaskToCurrentWeb(siteDesignId);
};
_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getSiteDesignRunStatus = function (runId) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["SiteDesigns"])(this, "").getSiteDesignRunStatus(undefined, runId);
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "eKJE":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/profiles/index.js ***!
  \************************************************/
/*! exports provided: Profiles, UrlZone */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "+QQQ");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Profiles", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Profiles"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UrlZone", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["UrlZone"]; });




Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_0__["SPRest"].prototype, "profiles", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_1__["Profiles"])(this._baseUrl);
    },
});
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "eVhm":
/*!*************************************************!*\
  !*** ./node_modules/@pnp/common/collections.js ***!
  \*************************************************/
/*! exports provided: objectToMap, mergeMaps */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "objectToMap", function() { return objectToMap; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "mergeMaps", function() { return mergeMaps; });
/* harmony import */ var _util__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./util */ "97IE");

/**
 * Used to calculate the object properties, with polyfill if needed
 */
var objectEntries = Object(_util__WEBPACK_IMPORTED_MODULE_0__["isFunc"])(Object.entries) ? Object.entries : function (o) { return Object.keys(o).map(function (k) { return [k, o[k]]; }); };
/**
 * Converts the supplied object to a map
 *
 * @param o The object to map
 */
function objectToMap(o) {
    if (o !== undefined && o !== null) {
        return new Map(objectEntries(o));
    }
    return new Map();
}
/**
 * Merges to Map instances together, overwriting values in target with matching keys, last in wins
 *
 * @param target map into which the other maps are merged
 * @param maps One or more maps to merge into the target
 */
function mergeMaps(target) {
    var maps = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        maps[_i - 1] = arguments[_i];
    }
    for (var i = 0; i < maps.length; i++) {
        maps[i].forEach(function (v, k) {
            target.set(k, v);
        });
    }
    return target;
}
//# sourceMappingURL=collections.js.map

/***/ }),

/***/ "fB9m":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/search/suggest.js ***!
  \************************************************/
/*! exports provided: _Suggest, Suggest */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Suggest", function() { return _Suggest; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Suggest", function() { return Suggest; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../telemetry */ "h3gW");





var _Suggest = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Suggest, _super);
    function _Suggest() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    _Suggest.prototype.execute = function (query) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var response, mapper;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.mapQueryToQueryString(query);
                        return [4 /*yield*/, this.get()];
                    case 1:
                        response = _a.sent();
                        mapper = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(response, "suggest") ? function (s_1) { return response.suggest[s_1].results; } : function (s_2) { return response[s_2]; };
                        return [2 /*return*/, {
                                PeopleNames: mapper("PeopleNames"),
                                PersonalResults: mapper("PersonalResults"),
                                Queries: mapper("Queries"),
                            }];
                }
            });
        });
    };
    _Suggest.prototype.mapQueryToQueryString = function (query) {
        var _this = this;
        var setProp = function (q) { return function (checkProp) { return function (sp) {
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_2__["hOP"])(q, checkProp)) {
                _this.query.set(sp, q[checkProp].toString());
            }
        }; }; };
        this.query.set("querytext", "'" + query.querytext + "'");
        var querySetter = setProp(query);
        querySetter("count")("inumberofquerysuggestions");
        querySetter("personalCount")("inumberofresultsuggestions");
        querySetter("preQuery")("fprequerysuggestions");
        querySetter("hitHighlighting")("fhithighlighting");
        querySetter("capitalize")("fcapitalizefirstletters");
        querySetter("culture")("culture");
        querySetter("stemming")("enablestemming");
        querySetter("includePeople")("showpeoplenamesuggestions");
        querySetter("queryRules")("enablequeryrules");
        querySetter("prefixMatch")("fprefixmatchallterms");
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_4__["tag"])("su.execute")
    ], _Suggest.prototype, "execute", null);
    _Suggest = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("_api/search/suggest")
    ], _Suggest);
    return _Suggest;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Suggest = function (baseUrl, options) {
    if (options === void 0) { options = {}; }
    return function (query) {
        return (new _Suggest(baseUrl)).configure(options).execute(query);
    };
};
//# sourceMappingURL=suggest.js.map

/***/ }),

/***/ "faye":
/*!****************************!*\
  !*** external "react-dom" ***!
  \****************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_faye__;

/***/ }),

/***/ "h3gW":
/*!*******************************************!*\
  !*** ./node_modules/@pnp/sp/telemetry.js ***!
  \*******************************************/
/*! exports provided: tag */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "tag", function() { return tag; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");



/**
 * Includes this method name in the X-ClientService-ClientTag used to record pnpjs usage
 *
 * @param name Method name, displayed in the
 */
function tag(name) {
    return function (target, key, descriptor) {
        if (descriptor === undefined) {
            descriptor = Object.getOwnPropertyDescriptor(target, key);
        }
        var originalMethod = descriptor.value;
        descriptor.value = function () {
            var args = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                args[_i] = arguments[_i];
            }
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
                return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                    this.configure(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "X-PnPjs-Tracking": name }));
                    return [2 /*return*/, originalMethod.apply(this, args)];
                });
            });
        };
        return descriptor;
    };
}
tag.getClientTag = function (h, deleteFromCollection) {
    if (deleteFromCollection === void 0) { deleteFromCollection = true; }
    if (h.has("X-PnPjs-Tracking")) {
        var methodName = h.get("X-PnPjs-Tracking");
        if (deleteFromCollection) {
            h.delete("X-PnPjs-Tracking");
        }
        if (!Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["stringIsNullOrEmpty"])(methodName)) {
            return methodName;
        }
    }
    return "";
};
tag.configure = function (o, name) {
    return o.configure(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "X-PnPjs-Tracking": name }));
};
tag.isTagged = function (o) {
    return o.data.options.headers && o.data.options.headers["X-PnPjs-Tracking"];
};
//# sourceMappingURL=telemetry.js.map

/***/ }),

/***/ "hMpi":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/decorators.js ***!
  \********************************************/
/*! exports provided: defaultPath */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "defaultPath", function() { return defaultPath; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/**
 * Class Decorators
 */

/**
 * Decorator used to specify the default path for SharePointQueryable objects
 *
 * @param path
 */
function defaultPath(path) {
    return function (target) {
        return /** @class */ (function (_super) {
            Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(class_1, _super);
            function class_1() {
                var args = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    args[_i] = arguments[_i];
                }
                return _super.call(this, args[0], args.length > 1 && args[1] !== undefined ? args[1] : path) || this;
            }
            return class_1;
        }(target));
    };
}
// TODO::?
// performance tracking method decorator
// redirect to graph api
//# sourceMappingURL=decorators.js.map

/***/ }),

/***/ "hy0S":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/lists/types.js ***!
  \*********************************************/
/*! exports provided: _Lists, Lists, _List, List, RenderListDataOptions, ControlMode */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Lists", function() { return _Lists; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Lists", function() { return Lists; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_List", function() { return _List; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "List", function() { return List; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RenderListDataOptions", function() { return RenderListDataOptions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ControlMode", function() { return ControlMode; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _utils_toResourcePath__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../utils/toResourcePath */ "t8hP");











var _Lists = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Lists, _super);
    function _Lists() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a list from the collection by guid id
     *
     * @param id The Id of the list (GUID)
     */
    _Lists.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(List(this).concat("('" + id + "')"), "ls.getById");
    };
    /**
     * Gets a list from the collection by title
     *
     * @param title The title of the list
     */
    _Lists.prototype.getByTitle = function (title) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(List(this, "getByTitle('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(title) + "')"), "ls.getByTitle");
    };
    /**
     * Adds a new list to the collection
     *
     * @param title The new list's title
     * @param description The new list's description
     * @param template The list template value
     * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
     * @param additionalSettings Will be passed as part of the list creation body
     */
    _Lists.prototype.add = function (title, desc, template, enableContentTypes, additionalSettings) {
        if (desc === void 0) { desc = ""; }
        if (template === void 0) { template = 100; }
        if (enableContentTypes === void 0) { enableContentTypes = false; }
        if (additionalSettings === void 0) { additionalSettings = {}; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var addSettings, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        addSettings = Object.assign({
                            "AllowContentTypes": enableContentTypes,
                            "BaseTemplate": template,
                            "ContentTypesEnabled": enableContentTypes,
                            "Description": desc,
                            "Title": title,
                        }, Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.List"), additionalSettings);
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])(addSettings))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, { data: data, list: this.getByTitle(addSettings.Title) }];
                }
            });
        });
    };
    /**
     * Ensures that the specified list exists in the collection (note: this method not supported for batching)
     *
     * @param title The new list's title
     * @param desc The new list's description
     * @param template The list template value
     * @param enableContentTypes If true content types will be allowed and enabled, otherwise they will be disallowed and not enabled
     * @param additionalSettings Will be passed as part of the list creation body or used to update an existing list
     */
    _Lists.prototype.ensure = function (title, desc, template, enableContentTypes, additionalSettings) {
        if (desc === void 0) { desc = ""; }
        if (template === void 0) { template = 100; }
        if (enableContentTypes === void 0) { enableContentTypes = false; }
        if (additionalSettings === void 0) { additionalSettings = {}; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var addOrUpdateSettings, list, data, e_1, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.hasBatch) {
                            throw Error("The ensure list method is not supported for use in a batch.");
                        }
                        addOrUpdateSettings = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(additionalSettings, { Title: title, Description: desc, ContentTypesEnabled: enableContentTypes }, true);
                        list = this.getByTitle(addOrUpdateSettings.Title);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 6]);
                        // this will throw if the list doesn't exist
                        return [4 /*yield*/, list.select("Title")()];
                    case 2:
                        // this will throw if the list doesn't exist
                        _a.sent();
                        return [4 /*yield*/, list.update(addOrUpdateSettings).then(function (r) { return r.data; })];
                    case 3:
                        data = _a.sent();
                        return [2 /*return*/, { created: false, data: data, list: this.getByTitle(addOrUpdateSettings.Title) }];
                    case 4:
                        e_1 = _a.sent();
                        return [4 /*yield*/, this.add(title, desc, template, enableContentTypes, addOrUpdateSettings).then(function (r) { return r.data; })];
                    case 5:
                        data = _a.sent();
                        return [2 /*return*/, { created: true, data: data, list: this.getByTitle(addOrUpdateSettings.Title) }];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Gets a list that is the default asset location for images or other files, which the users upload to their wiki pages.
     */
    _Lists.prototype.ensureSiteAssetsLibrary = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var json;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Lists, "ensuresiteassetslibrary"))];
                    case 1:
                        json = _a.sent();
                        return [2 /*return*/, List(Object(_odata__WEBPACK_IMPORTED_MODULE_4__["odataUrlFrom"])(json))];
                }
            });
        });
    };
    /**
     * Gets a list that is the default location for wiki pages.
     */
    _Lists.prototype.ensureSitePagesLibrary = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var json;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Lists, "ensuresitepageslibrary"))];
                    case 1:
                        json = _a.sent();
                        return [2 /*return*/, List(Object(_odata__WEBPACK_IMPORTED_MODULE_4__["odataUrlFrom"])(json))];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("ls.add")
    ], _Lists.prototype, "add", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("ls.ensure")
    ], _Lists.prototype, "ensure", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("ls.ensureSiteAssetsLibrary")
    ], _Lists.prototype, "ensureSiteAssetsLibrary", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("ls.ensureSitePagesLibrary")
    ], _Lists.prototype, "ensureSitePagesLibrary", null);
    _Lists = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("lists")
    ], _Lists);
    return _Lists;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableCollection"]));

var Lists = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_Lists);
var _List = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_List, _super);
    function _List() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["deleteableWithETag"])("l");
        return _this;
    }
    Object.defineProperty(_List.prototype, "effectiveBasePermissions", {
        /**
         * Gets the effective base permissions of this list
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryable"])(this, "EffectiveBasePermissions"), "l.effectiveBasePermissions");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_List.prototype, "eventReceivers", {
        /**
         * Gets the event receivers attached to this list
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryableCollection"])(this, "EventReceivers"), "l.eventReceivers");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_List.prototype, "relatedFields", {
        /**
         * Gets the related fields of this list
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryable"])(this, "getRelatedFields"), "l.relatedFields");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_List.prototype, "informationRightsManagementSettings", {
        /**
         * Gets the IRM settings for this list
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["SharePointQueryable"])(this, "InformationRightsManagementSettings"), "l.informationRightsManagementSettings");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Updates this list intance with the supplied properties
     *
     * @param properties A plain object hash of values to update for the list
     * @param eTag Value used in the IF-Match header, by default "*"
     */
    _List.prototype.update = function (properties, eTag) {
        if (eTag === void 0) { eTag = "*"; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, data, list;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.List"), properties), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({
                            "IF-Match": eTag,
                            "X-HTTP-Method": "MERGE",
                        }));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this, postBody)];
                    case 1:
                        data = _a.sent();
                        list = this;
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(properties, "Title")) {
                            list = this.getParent(List, this.parentUrl, "getByTitle('" + properties.Title + "')");
                        }
                        return [2 /*return*/, {
                                data: data,
                                list: list,
                            }];
                }
            });
        });
    };
    /**
     * Returns the collection of changes from the change log that have occurred within the list, based on the specified query.
     * @param query A query that is performed against the change log.
     */
    _List.prototype.getChanges = function (query) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(List, "getchanges"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({ query: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.ChangeQuery"), query) }));
    };
    /**
     * Returns the collection of items in the list based on the provided CamlQuery
     * @param query A query that is performed against the list
     * @param expands An expanded array of n items that contains fields to expand in the CamlQuery
     */
    _List.prototype.getItemsByCAMLQuery = function (query) {
        var expands = [];
        for (var _i = 1; _i < arguments.length; _i++) {
            expands[_i - 1] = arguments[_i];
        }
        var q = this.clone(List, "getitems");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(q.expand.apply(q, expands), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({ query: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.CamlQuery"), query) }));
    };
    /**
     * See: https://msdn.microsoft.com/en-us/library/office/dn292554.aspx
     * @param query An object that defines the change log item query
     */
    _List.prototype.getListItemChangesSinceToken = function (query) {
        var o = this.clone(List, "getlistitemchangessincetoken").usingParser({ parse: function (r) { return r.text(); } });
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(o, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({ "query": Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.ChangeLogItemQuery"), query) }));
    };
    /**
     * Moves the list to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     */
    _List.prototype.recycle = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(List, "recycle"))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(data, "Recycle") ? data.Recycle : data];
                }
            });
        });
    };
    /**
     * Renders list data based on the view xml provided
     * @param viewXml A string object representing a view xml
     */
    _List.prototype.renderListData = function (viewXml) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var q, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        q = this.clone(List, "renderlistdata(@viewXml)");
                        q.query.set("@viewXml", "'" + viewXml + "'");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(q)];
                    case 1:
                        data = _a.sent();
                        // data will be a string, so we parse it again
                        return [2 /*return*/, JSON.parse(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(data, "RenderListData") ? data.RenderListData : data)];
                }
            });
        });
    };
    /**
     * Returns the data for the specified query view
     *
     * @param parameters The parameters to be used to render list data as JSON string.
     * @param overrideParameters The parameters that are used to override and extend the regular SPRenderListDataParameters.
     * @param queryParams Allows setting of query parameters
     */
    _List.prototype.renderListDataAsStream = function (parameters, overrideParameters, queryParams) {
        if (overrideParameters === void 0) { overrideParameters = null; }
        if (queryParams === void 0) { queryParams = new Map(); }
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(parameters, "RenderOptions") && Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isArray"])(parameters.RenderOptions)) {
            parameters.RenderOptions = parameters.RenderOptions.reduce(function (v, c) { return v + c; });
        }
        var postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({
            overrideParameters: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.RenderListDataOverrideParameters"), overrideParameters),
            parameters: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.RenderListDataParameters"), parameters),
        });
        var clone = this.clone(List, "RenderListDataAsStream", true);
        if (queryParams && queryParams.size > 0) {
            queryParams.forEach(function (v, k) { return clone.query.set(k, v); });
        }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(clone, postBody);
    };
    /**
     * Gets the field values and field schema attributes for a list item.
     * @param itemId Item id of the item to render form data for
     * @param formId The id of the form
     * @param mode Enum representing the control mode of the form (Display, Edit, New)
     */
    _List.prototype.renderListFormData = function (itemId, formId, mode) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(List, "renderlistformdata(itemid=" + itemId + ", formid='" + formId + "', mode='" + mode + "')"))];
                    case 1:
                        data = _a.sent();
                        // data will be a string, so we parse it again
                        return [2 /*return*/, JSON.parse(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(data, "RenderListFormData") ? data.RenderListFormData : data)];
                }
            });
        });
    };
    /**
     * Reserves a list item ID for idempotent list item creation.
     */
    _List.prototype.reserveListItemId = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(List, "reservelistitemid"))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(data, "ReserveListItemId") ? data.ReserveListItemId : data];
                }
            });
        });
    };
    /**
     * Returns the ListItemEntityTypeFullName for this list, used when adding/updating list items. Does not support batching.
     */
    _List.prototype.getListItemEntityTypeFullName = function () {
        return this.clone(List, null, false).select("ListItemEntityTypeFullName").get().then(function (o) { return o.ListItemEntityTypeFullName; });
    };
    /**
     * Creates an item using path (in a folder), validates and sets its field values.
     *
     * @param formValues The fields to change and their new values.
     * @param decodedUrl Path decoded url; folder's server relative path.
     * @param bNewDocumentUpdate true if the list item is a document being updated after upload; otherwise false.
     * @param checkInComment Optional check in comment.
     * @param additionalProps Optional set of additional properties LeafName new document file name,
     */
    _List.prototype.addValidateUpdateItemUsingPath = function (formValues, decodedUrl, bNewDocumentUpdate, checkInComment, additionalProps) {
        if (bNewDocumentUpdate === void 0) { bNewDocumentUpdate = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var addProps, res;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        addProps = {
                            FolderPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_10__["toResourcePath"])(decodedUrl),
                        };
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["objectDefinedNotNull"])(additionalProps)) {
                            if (additionalProps.leafName) {
                                addProps.LeafName = Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_10__["toResourcePath"])(additionalProps.leafName);
                            }
                            if (additionalProps.objectType) {
                                addProps.UnderlyingObjectType = additionalProps.objectType;
                            }
                        }
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(List, "AddValidateUpdateItemUsingPath()"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({
                                bNewDocumentUpdate: bNewDocumentUpdate,
                                checkInComment: checkInComment,
                                formValues: formValues,
                                listItemCreateInfo: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("SP.ListItemCreationInformationUsingPath"), addProps),
                            }))];
                    case 1:
                        res = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(res, "AddValidateUpdateItemUsingPath") ? res.AddValidateUpdateItemUsingPath : res];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.update")
    ], _List.prototype, "update", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.getChanges")
    ], _List.prototype, "getChanges", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.CAMLQuery")
    ], _List.prototype, "getItemsByCAMLQuery", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.ChangesSinceToken")
    ], _List.prototype, "getListItemChangesSinceToken", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.recycle")
    ], _List.prototype, "recycle", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.renderListData")
    ], _List.prototype, "renderListData", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.AsStream")
    ], _List.prototype, "renderListDataAsStream", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.renderListFormData")
    ], _List.prototype, "renderListFormData", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.reserveListItemId")
    ], _List.prototype, "reserveListItemId", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.getListItemEntityTypeFullName")
    ], _List.prototype, "getListItemEntityTypeFullName", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_9__["tag"])("l.addValidateUpdateItemUsingPath")
    ], _List.prototype, "addValidateUpdateItemUsingPath", null);
    return _List;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["_SharePointQueryableInstance"]));

var List = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_3__["spInvokableFactory"])(_List);
/**
 * Enum representing the options of the RenderOptions property on IRenderListDataParameters interface
 */
var RenderListDataOptions;
(function (RenderListDataOptions) {
    RenderListDataOptions[RenderListDataOptions["None"] = 0] = "None";
    RenderListDataOptions[RenderListDataOptions["ContextInfo"] = 1] = "ContextInfo";
    RenderListDataOptions[RenderListDataOptions["ListData"] = 2] = "ListData";
    RenderListDataOptions[RenderListDataOptions["ListSchema"] = 4] = "ListSchema";
    RenderListDataOptions[RenderListDataOptions["MenuView"] = 8] = "MenuView";
    RenderListDataOptions[RenderListDataOptions["ListContentType"] = 16] = "ListContentType";
    RenderListDataOptions[RenderListDataOptions["FileSystemItemId"] = 32] = "FileSystemItemId";
    RenderListDataOptions[RenderListDataOptions["ClientFormSchema"] = 64] = "ClientFormSchema";
    RenderListDataOptions[RenderListDataOptions["QuickLaunch"] = 128] = "QuickLaunch";
    RenderListDataOptions[RenderListDataOptions["Spotlight"] = 256] = "Spotlight";
    RenderListDataOptions[RenderListDataOptions["Visualization"] = 512] = "Visualization";
    RenderListDataOptions[RenderListDataOptions["ViewMetadata"] = 1024] = "ViewMetadata";
    RenderListDataOptions[RenderListDataOptions["DisableAutoHyperlink"] = 2048] = "DisableAutoHyperlink";
    RenderListDataOptions[RenderListDataOptions["EnableMediaTAUrls"] = 4096] = "EnableMediaTAUrls";
    RenderListDataOptions[RenderListDataOptions["ParentInfo"] = 8192] = "ParentInfo";
    RenderListDataOptions[RenderListDataOptions["PageContextInfo"] = 16384] = "PageContextInfo";
    RenderListDataOptions[RenderListDataOptions["ClientSideComponentManifest"] = 32768] = "ClientSideComponentManifest";
})(RenderListDataOptions || (RenderListDataOptions = {}));
/**
 * Determines the display mode of the given control or view
 */
var ControlMode;
(function (ControlMode) {
    ControlMode[ControlMode["Display"] = 1] = "Display";
    ControlMode[ControlMode["Edit"] = 2] = "Edit";
    ControlMode[ControlMode["New"] = 3] = "New";
})(ControlMode || (ControlMode = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "hyGn":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/common/libconfig.js ***!
  \***********************************************/
/*! exports provided: setup, RuntimeConfigImpl, RuntimeConfig */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "setup", function() { return setup; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RuntimeConfigImpl", function() { return RuntimeConfigImpl; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RuntimeConfig", function() { return RuntimeConfig; });
/* harmony import */ var _collections__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./collections */ "eVhm");

function setup(config) {
    RuntimeConfig.assign(config);
}
// lable mapping for known config values
var s = [
    "defaultCachingStore",
    "defaultCachingTimeoutSeconds",
    "globalCacheDisable",
    "enableCacheExpiration",
    "cacheExpirationIntervalMilliseconds",
    "spfxContext",
    "ie11",
];
var RuntimeConfigImpl = /** @class */ (function () {
    function RuntimeConfigImpl(_v) {
        if (_v === void 0) { _v = new Map(); }
        this._v = _v;
        // setup defaults
        this._v.set(s[0], "session");
        this._v.set(s[1], 60);
        this._v.set(s[2], false);
        this._v.set(s[3], false);
        this._v.set(s[4], 750);
        this._v.set(s[5], null);
        this._v.set(s[6], false);
    }
    /**
     *
     * @param config The set of properties to add to the globa configuration instance
     */
    RuntimeConfigImpl.prototype.assign = function (config) {
        this._v = Object(_collections__WEBPACK_IMPORTED_MODULE_0__["mergeMaps"])(this._v, Object(_collections__WEBPACK_IMPORTED_MODULE_0__["objectToMap"])(config));
    };
    RuntimeConfigImpl.prototype.get = function (key) {
        return this._v.get(key);
    };
    Object.defineProperty(RuntimeConfigImpl.prototype, "defaultCachingStore", {
        get: function () {
            return this.get(s[0]);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(RuntimeConfigImpl.prototype, "defaultCachingTimeoutSeconds", {
        get: function () {
            return this.get(s[1]);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(RuntimeConfigImpl.prototype, "globalCacheDisable", {
        get: function () {
            return this.get(s[2]);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(RuntimeConfigImpl.prototype, "enableCacheExpiration", {
        get: function () {
            return this.get(s[3]);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(RuntimeConfigImpl.prototype, "cacheExpirationIntervalMilliseconds", {
        get: function () {
            return this.get(s[4]);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(RuntimeConfigImpl.prototype, "spfxContext", {
        get: function () {
            return this.get(s[5]);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(RuntimeConfigImpl.prototype, "ie11", {
        get: function () {
            return this.get(s[6]);
        },
        enumerable: true,
        configurable: true
    });
    return RuntimeConfigImpl;
}());

var _runtimeConfig = new RuntimeConfigImpl();
var RuntimeConfig = _runtimeConfig;
//# sourceMappingURL=libconfig.js.map

/***/ }),

/***/ "iV2g":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/views/list.js ***!
  \********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "32VY");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "views", _types__WEBPACK_IMPORTED_MODULE_2__["Views"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "defaultView", _types__WEBPACK_IMPORTED_MODULE_2__["View"], "DefaultView");
_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"].prototype.getView = function (viewId) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_2__["View"])(this, "getView('" + viewId + "')");
};
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "ikbv":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/index.js ***!
  \***********************************************/
/*! exports provided: SPSharedObjectType, SharingDomainRestrictionMode, SharingLinkKind, SharingOperationStatusCode, SharingRole, RoleType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _file__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./file */ "qPin");
/* harmony import */ var _folder__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./folder */ "VQtT");
/* harmony import */ var _item__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./item */ "K1m8");
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./web */ "8gyt");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./types */ "AvPO");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPSharedObjectType", function() { return _types__WEBPACK_IMPORTED_MODULE_4__["SPSharedObjectType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharingDomainRestrictionMode", function() { return _types__WEBPACK_IMPORTED_MODULE_4__["SharingDomainRestrictionMode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharingLinkKind", function() { return _types__WEBPACK_IMPORTED_MODULE_4__["SharingLinkKind"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharingOperationStatusCode", function() { return _types__WEBPACK_IMPORTED_MODULE_4__["SharingOperationStatusCode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SharingRole", function() { return _types__WEBPACK_IMPORTED_MODULE_4__["SharingRole"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RoleType", function() { return _types__WEBPACK_IMPORTED_MODULE_4__["RoleType"]; });






//# sourceMappingURL=index.js.map

/***/ }),

/***/ "ilAS":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/sites/types.js ***!
  \*********************************************/
/*! exports provided: _Site, Site */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Site", function() { return _Site; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Site", function() { return Site; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _batch__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../batch */ "8xxP");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _splibconfig__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! ../splibconfig */ "0p/9");














var _Site = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Site, _super);
    function _Site() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(_Site.prototype, "rootWeb", {
        /**
        * Gets the root web of the site collection
        *
        */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_webs_types__WEBPACK_IMPORTED_MODULE_3__["Web"])(this, "rootweb"), "si.rootWeb");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Gets a Web instance representing the root web of the site collection
     * correctly setup for chaining within the library
     */
    _Site.prototype.getRootWeb = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var web;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.rootWeb.select("Url")()];
                    case 1:
                        web = _a.sent();
                        return [2 /*return*/, _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_webs_types__WEBPACK_IMPORTED_MODULE_3__["Web"])(web.Url), "si.getRootWeb")];
                }
            });
        });
    };
    /**
    * Gets the context information for this site collection
    */
    _Site.prototype.getContextInfo = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var q, data, info;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        q = _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Site(this.parentUrl, "_api/contextinfo"), "si.getContextInfo");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(q)];
                    case 1:
                        data = _a.sent();
                        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["hOP"])(data, "GetContextWebInformation")) {
                            info = data.GetContextWebInformation;
                            info.SupportedSchemaVersions = info.SupportedSchemaVersions.results;
                            return [2 /*return*/, info];
                        }
                        else {
                            return [2 /*return*/, data];
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    _Site.prototype.createBatch = function () {
        return new _batch__WEBPACK_IMPORTED_MODULE_8__["SPBatch"](this.parentUrl);
    };
    /**
    * Deletes the current site
    *
    */
    _Site.prototype.delete = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var site, q;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(Site, "").select("Id")()];
                    case 1:
                        site = _a.sent();
                        q = _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Site(this.parentUrl, "_api/SPSiteManager/Delete"), "si.delete");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(q, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({ siteId: site.Id }))];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Gets the document libraries on a site. Static method. (SharePoint Online only)
     *
     * @param absoluteWebUrl The absolute url of the web whose document libraries should be returned
     */
    _Site.prototype.getDocumentLibraries = function (absoluteWebUrl) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var q, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        q = _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryable"])("", "_api/sp.web.getdocumentlibraries(@v)"), "si.getDocumentLibraries");
                        q.query.set("@v", "'" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_9__["escapeQueryStrValue"])(absoluteWebUrl) + "'");
                        return [4 /*yield*/, q()];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["hOP"])(data, "GetDocumentLibraries") ? data.GetDocumentLibraries : data];
                }
            });
        });
    };
    /**
     * Gets the site url from a page url
     *
     * @param absolutePageUrl The absolute url of the page
     */
    _Site.prototype.getWebUrlFromPageUrl = function (absolutePageUrl) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var q, data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        q = _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryable"])("", "_api/sp.web.getweburlfrompageurl(@v)"), "si.getWebUrlFromPageUrl");
                        q.query.set("@v", "'" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_9__["escapeQueryStrValue"])(absolutePageUrl) + "'");
                        return [4 /*yield*/, q()];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["hOP"])(data, "GetWebUrlFromPageUrl") ? data.GetWebUrlFromPageUrl : data];
                }
            });
        });
    };
    /**
     * Returns the collection of changes from the change log that have occurred within the list, based on the specified query
     *
     * @param query The change query
     */
    _Site.prototype.getChanges = function (query) {
        var postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({ "query": Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_11__["metadata"])("SP.ChangeQuery"), query) });
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(_webs_types__WEBPACK_IMPORTED_MODULE_3__["Web"], "getchanges"), postBody);
    };
    /**
    * Opens a web by id (using POST)
    *
    * @param webId The GUID id of the web to open
    */
    _Site.prototype.openWebById = function (webId) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var data;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Site, "openWebById('" + webId + "')"))];
                    case 1:
                        data = _a.sent();
                        return [2 /*return*/, {
                                data: data,
                                web: Object(_webs_types__WEBPACK_IMPORTED_MODULE_3__["Web"])(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_12__["extractWebUrl"])(Object(_odata__WEBPACK_IMPORTED_MODULE_6__["odataUrlFrom"])(data))),
                            }];
                }
            });
        });
    };
    /**
     * Creates a Modern communication site.
     *
     * @param title The title of the site to create
     * @param lcid The language to use for the site. If not specified will default to 1033 (English).
     * @param shareByEmailEnabled If set to true, it will enable sharing files via Email. By default it is set to false
     * @param url The fully qualified URL (e.g. https://yourtenant.sharepoint.com/sites/mysitecollection) of the site.
     * @param description The description of the communication site.
     * @param classification The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
     * @param siteDesignId The Guid of the site design to be used.
     *                     You can use the below default OOTB GUIDs:
     *                     Topic: 00000000-0000-0000-0000-000000000000
     *                     Showcase: 6142d2a0-63a5-4ba0-aede-d9fefca2c767
     *                     Blank: f6cc5403-0d63-442e-96c0-285923709ffc
     * @param hubSiteId The id of the hub site to which the new site should be associated
     * @param owner Optional owner value, required if executing the method in app only mode
     */
    _Site.prototype.createCommunicationSite = function (title, lcid, shareByEmailEnabled, url, description, classification, siteDesignId, hubSiteId, owner) {
        if (lcid === void 0) { lcid = 1033; }
        if (shareByEmailEnabled === void 0) { shareByEmailEnabled = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                return [2 /*return*/, this.createCommunicationSiteFromProps({
                        Classification: classification,
                        Description: description,
                        HubSiteId: hubSiteId,
                        Lcid: lcid,
                        Owner: owner,
                        ShareByEmailEnabled: shareByEmailEnabled,
                        SiteDesignId: siteDesignId,
                        Title: title,
                        Url: url,
                    })];
            });
        });
    };
    _Site.prototype.createCommunicationSiteFromProps = function (props) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var p, postBody;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                p = Object.assign({}, {
                    Classification: "",
                    Description: "",
                    HubSiteId: _splibconfig__WEBPACK_IMPORTED_MODULE_13__["emptyGuid"],
                    Lcid: 1033,
                    ShareByEmailEnabled: false,
                    SiteDesignId: _splibconfig__WEBPACK_IMPORTED_MODULE_13__["emptyGuid"],
                    WebTemplate: "SITEPAGEPUBLISHING#0",
                    WebTemplateExtensionId: _splibconfig__WEBPACK_IMPORTED_MODULE_13__["emptyGuid"],
                }, props);
                postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({
                    "request": Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_11__["metadata"])("Microsoft.SharePoint.Portal.SPSiteCreationRequest"), p),
                });
                return [2 /*return*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Site(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_12__["extractWebUrl"])(this.toUrl()), "/_api/SPSiteManager/Create"), postBody)];
            });
        });
    };
    /**
     *
     * @param url Site Url that you want to check if exists
     */
    _Site.prototype.exists = function (url) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, value;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])({ url: url });
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Site(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_12__["extractWebUrl"])(this.toUrl()), "/_api/SP.Site.Exists"), postBody)];
                    case 1:
                        value = _a.sent();
                        return [2 /*return*/, value];
                }
            });
        });
    };
    /**
    * Creates a Modern team site backed by Office 365 group. For use in SP Online only. This will not work with App-only tokens
    *
    * @param displayName The title or display name of the Modern team site to be created
    * @param alias Alias of the underlying Office 365 Group
    * @param isPublic Defines whether the Office 365 Group will be public (default), or private.
    * @param lcid The language to use for the site. If not specified will default to English (1033).
    * @param description The description of the site to be created.
    * @param classification The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
    * @param owners The Owners of the site to be created
    */
    _Site.prototype.createModernTeamSite = function (displayName, alias, isPublic, lcid, description, classification, owners, hubSiteId, siteDesignId) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                return [2 /*return*/, this.createModernTeamSiteFromProps({
                        alias: alias,
                        classification: classification,
                        description: description,
                        displayName: displayName,
                        hubSiteId: hubSiteId,
                        isPublic: isPublic,
                        lcid: lcid,
                        owners: owners,
                        siteDesignId: siteDesignId,
                    })];
            });
        });
    };
    _Site.prototype.createModernTeamSiteFromProps = function (props) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var p, postBody;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                p = Object.assign({}, {
                    classification: "",
                    description: "",
                    hubSiteId: _splibconfig__WEBPACK_IMPORTED_MODULE_13__["emptyGuid"],
                    isPublic: true,
                    lcid: 1033,
                    owners: [],
                }, props);
                postBody = {
                    alias: p.alias,
                    displayName: p.displayName,
                    isPublic: p.isPublic,
                    optionalParams: {
                        Classification: p.classification,
                        CreationOptions: {
                            "results": ["SPSiteLanguage:" + p.lcid, "HubSiteId:" + p.hubSiteId],
                        },
                        Description: p.description,
                        Owners: {
                            "results": p.owners,
                        },
                    },
                };
                if (p.siteDesignId) {
                    postBody.optionalParams.CreationOptions.results.push("implicit_formula_292aa8a00786498a87a5ca52d9f4214a_" + p.siteDesignId);
                }
                return [2 /*return*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Site(Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_12__["extractWebUrl"])(this.toUrl()), "/_api/GroupSiteManager/CreateGroupEx"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])(postBody))];
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("si.getChanges")
    ], _Site.prototype, "getChanges", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("si.openWebById")
    ], _Site.prototype, "openWebById", null);
    _Site = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_2__["defaultPath"])("_api/site")
    ], _Site);
    return _Site;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Site = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Site);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "jM62":
/*!******************************************************!*\
  !*** ./node_modules/@pnp/sp/column-defaults/list.js ***!
  \******************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _folders_types__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../folders/types */ "4tXH");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");
/* harmony import */ var _pnp_logging__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @pnp/logging */ "B3L+");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _presets_all__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../presets/all */ "F/m7");









Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_2__["_List"], "rootFolder", _folders_types__WEBPACK_IMPORTED_MODULE_3__["Folder"], "rootFolder");
_lists_types__WEBPACK_IMPORTED_MODULE_2__["_List"].prototype.getDefaultColumnValues = function () {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var pathPart, webUrl, path, baseFilePath, xml, e_1, matches, tags;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, this.rootFolder.select("ServerRelativePath")()];
                case 1:
                    pathPart = _a.sent();
                    return [4 /*yield*/, this.select("ParentWeb/Url").expand("ParentWeb")()];
                case 2:
                    webUrl = _a.sent();
                    path = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["combine"])("/", pathPart.ServerRelativePath.DecodedUrl, "Forms/client_LocationBasedDefaults.html");
                    baseFilePath = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["combine"])(webUrl.ParentWeb.Url, "_api/web", "getFileByServerRelativePath(decodedUrl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_5__["escapeQueryStrValue"])(path) + "')");
                    xml = "";
                    _a.label = 3;
                case 3:
                    _a.trys.push([3, 5, , 6]);
                    return [4 /*yield*/, Object(_folders_types__WEBPACK_IMPORTED_MODULE_3__["Folder"])(baseFilePath, "$value").usingParser(new _pnp_odata__WEBPACK_IMPORTED_MODULE_1__["TextParser"]())(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["headers"])({ "binaryStringResponseBody": "true" }))];
                case 4:
                    xml = _a.sent();
                    return [3 /*break*/, 6];
                case 5:
                    e_1 = _a.sent();
                    // if this call fails we assume it is because the file is 404
                    if (e_1 && e_1.status && e_1.status === 404) {
                        // return an empty array
                        return [2 /*return*/, []];
                    }
                    throw e_1;
                case 6:
                    matches = xml.match(/<a.*?<\/a>/ig);
                    tags = matches === null ? [] : matches.map(function (t) { return t.trim(); });
                    // now we need to turn these tags of form into objects
                    // <a href="/sites/dev/My%20Title"><DefaultValue FieldName="TextField">Test</DefaultValue></a>
                    return [2 /*return*/, tags.map(function (t) {
                            var m = /<a href="(.*?)"><DefaultValue FieldName="(.*?)">(.*?)<\/DefaultValue>/ig.exec(t);
                            // if things worked our captures are:
                            // 0: whole string
                            // 1: ENCODED server relative path
                            // 2: Field internal name
                            // 3: Default value as string
                            if (m.length < 1) {
                                // this indicates an error somewhere, but we have no way to meaningfully recover
                                // perhaps the way the tags are stored has changed on the server? Check that first.
                                _pnp_logging__WEBPACK_IMPORTED_MODULE_6__["Logger"].write("Could not parse default column value from '" + t + "'", 2 /* Warning */);
                                return null;
                            }
                            // return the parsed out values
                            return {
                                name: m[2],
                                path: decodeURIComponent(m[1]),
                                value: m[3],
                            };
                        }).filter(function (v) { return v !== null; })];
            }
        });
    });
};
_lists_types__WEBPACK_IMPORTED_MODULE_2__["_List"].prototype.setDefaultColumnValues = function (defaults) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var fieldDefs, tags, xml, pathPart, webUrl, path, baseFilePath, existingReceivers;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_presets_all__WEBPACK_IMPORTED_MODULE_8__["SharePointQueryableCollection"])(this, "fields").select("InternalName", "TypeAsString").filter("Hidden ne true")()];
                case 1:
                    fieldDefs = _a.sent();
                    tags = defaults.map(function (fieldDefault) {
                        var index = fieldDefs.findIndex(function (fd) { return fd.InternalName === fieldDefault.name; });
                        if (index < 0) {
                            throw Error("Field '" + fieldDefault.name + "' does not exist in the list. Please check the internal field name. Failed to set defaults.");
                        }
                        var fieldDef = fieldDefs[index];
                        var value = "";
                        switch (fieldDef.TypeAsString) {
                            case "Boolean":
                            case "Currency":
                            case "Text":
                            case "DateTime":
                            case "Number":
                            case "Choice":
                            case "User":
                                if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["isArray"])(fieldDefault.value)) {
                                    throw Error("The type '" + fieldDef.TypeAsString + "' does not support multiple values.");
                                }
                                value = "" + fieldDefault.value;
                                break;
                            case "MultiChoice":
                                if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["isArray"])(fieldDefault.value)) {
                                    value = fieldDefault.value.map(function (v) { return "" + v; }).join(";");
                                }
                                else {
                                    value = "" + fieldDefault.value;
                                }
                                break;
                            case "UserMulti":
                                if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["isArray"])(fieldDefault.value)) {
                                    value = fieldDefault.value.map(function (v) { return "" + v; }).join(";#");
                                }
                                else {
                                    value = "" + fieldDefault.value;
                                }
                                break;
                            case "Taxonomy":
                                if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["isArray"])(fieldDefault.value)) {
                                    throw Error("The type '" + fieldDef.TypeAsString + "' does not support multiple values.");
                                }
                                else {
                                    value = fieldDefault.value.wssId + ";#" + fieldDefault.value.termName + "|" + fieldDefault.value.termId;
                                }
                                break;
                            case "TaxonomyMulti":
                                if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["isArray"])(fieldDefault.value)) {
                                    value = fieldDefault.value.map(function (v) { return v.wssId + ";#" + v.termName + "|" + v.termId; }).join(";#");
                                }
                                value = fieldDefault.value.wssId + ";#" + fieldDefault.value.termName + "|" + fieldDefault.value.termId;
                                break;
                        }
                        return "<a href=\"" + fieldDefault.path.replace(/ /gi, "%20") + "\"><DefaultValue FieldName=\"" + fieldDefault.name + "\">" + value + "</DefaultValue></a>";
                    });
                    xml = "<MetadataDefaults>" + tags.join("") + "</MetadataDefaults>";
                    return [4 /*yield*/, this.rootFolder.select("ServerRelativePath")()];
                case 2:
                    pathPart = _a.sent();
                    return [4 /*yield*/, this.select("ParentWeb/Url").expand("ParentWeb")()];
                case 3:
                    webUrl = _a.sent();
                    path = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["combine"])("/", pathPart.ServerRelativePath.DecodedUrl, "Forms");
                    baseFilePath = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_4__["combine"])(webUrl.ParentWeb.Url, "_api/web", "getFolderByServerRelativePath(decodedUrl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_5__["escapeQueryStrValue"])(path) + "')", "files");
                    return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Object(_folders_types__WEBPACK_IMPORTED_MODULE_3__["Folder"])(baseFilePath, "add(overwrite=true,url='client_LocationBasedDefaults.html')"), { body: xml })];
                case 4:
                    _a.sent();
                    return [4 /*yield*/, this.eventReceivers.filter("ReceiverName eq 'LocationBasedMetadataDefaultsReceiver ItemAdded'").select("ReceiverId")()];
                case 5:
                    existingReceivers = _a.sent();
                    if (!(existingReceivers.length < 1)) return [3 /*break*/, 7];
                    return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Object(_lists_types__WEBPACK_IMPORTED_MODULE_2__["List"])(this.eventReceivers, "add"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_1__["body"])({
                            eventReceiverCreationInformation: {
                                EventType: 10001,
                                ReceiverAssembly: "Microsoft.Office.DocumentManagement, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
                                ReceiverClass: "Microsoft.Office.DocumentManagement.LocationBasedMetadataDefaultsReceiver",
                                ReceiverName: "LocationBasedMetadataDefaultsReceiver ItemAdded",
                                SequenceNumber: 1000,
                                Synchronization: 1,
                            },
                        }))];
                case 6:
                    _a.sent();
                    _a.label = 7;
                case 7: return [2 /*return*/];
            }
        });
    });
};
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "lYrR":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/items/index.js ***!
  \*********************************************/
/*! exports provided: Item, Items, ItemVersion, ItemVersions, PagedItemCollection */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list */ "NTTg");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "3DT9");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Item", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Item"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Items", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["Items"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ItemVersion", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ItemVersion"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ItemVersions", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["ItemVersions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PagedItemCollection", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["PagedItemCollection"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "laIm":
/*!******************************************************!*\
  !*** ./node_modules/@pnp/sp/clientside-pages/web.js ***!
  \******************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "UPqE");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");



_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.getClientsideWebParts = function () {
    return this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableCollection"], "GetClientSideWebParts")();
};
_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.addClientsidePage =
    function (pageName, title, layout, promotedState) {
        if (title === void 0) { title = pageName.replace(/\.[^/.]+$/, ""); }
        return Object(_types__WEBPACK_IMPORTED_MODULE_1__["CreateClientsidePage"])(this, pageName, title, layout, promotedState);
    };
_webs_types__WEBPACK_IMPORTED_MODULE_0__["_Web"].prototype.loadClientsidePage = function (path) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["ClientsidePageFromFile"])(this.getFileByServerRelativePath(path));
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "ld+d":
/*!********************************************************!*\
  !*** ./node_modules/@pnp/sp/clientside-pages/funcs.js ***!
  \********************************************************/
/*! exports provided: getNextOrder, reindex */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getNextOrder", function() { return getNextOrder; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "reindex", function() { return reindex; });
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/common */ "zqEK");

/**
 * Gets the next order value 1 based for the provided collection
 *
 * @param collection Collection of orderable things
 */
function getNextOrder(collection) {
    return collection.length < 1 ? 1 : (Math.max.apply(null, collection.map(function (i) { return i.order; })) + 1);
}
/**
 * Normalizes the order value for all the sections, columns, and controls to be 1 based and stepped (1, 2, 3...)
 *
 * @param collection The collection to normalize
 */
function reindex(collection) {
    for (var i = 0; i < collection.length; i++) {
        collection[i].order = i + 1;
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["hOP"])(collection[i], "columns")) {
            reindex(collection[i].columns);
        }
        else if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["hOP"])(collection[i], "controls")) {
            reindex(collection[i].controls);
        }
    }
}
//# sourceMappingURL=funcs.js.map

/***/ }),

/***/ "luMW":
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/types.js ***!
  \*********************************************************/
/*! exports provided: _RegionalSettings, RegionalSettings, _TimeZone, TimeZone, _TimeZones, TimeZones */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_RegionalSettings", function() { return _RegionalSettings; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "RegionalSettings", function() { return RegionalSettings; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_TimeZone", function() { return _TimeZone; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TimeZone", function() { return TimeZone; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_TimeZones", function() { return _TimeZones; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TimeZones", function() { return TimeZones; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../telemetry */ "h3gW");







var _RegionalSettings = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_RegionalSettings, _super);
    function _RegionalSettings() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(_RegionalSettings.prototype, "installedLanguages", {
        /**
         * Gets the collection of languages used in a server farm.
         * ** Please use getInstalledLanguages instead of this method **
         */
        get: function () {
            console.warn("Deprecated: RegionalSettings.installedLanguages is deprecated, please use RegionalSettings.getInstalledLanguages");
            return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableCollection"])(this, "installedlanguages"), "rs.installedLanguages");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_RegionalSettings.prototype, "timeZone", {
        /**
         * Gets time zone
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(TimeZone(this), "rs.tz");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_RegionalSettings.prototype, "timeZones", {
        /**
         * Gets time zones
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(TimeZones(this), "rs.tzs");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Gets the collection of languages used in a server farm.
     */
    _RegionalSettings.prototype.getInstalledLanguages = function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var results;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableCollection"])(this, "installedlanguages"), "rs.getInstalledLanguages")()];
                    case 1:
                        results = _a.sent();
                        return [2 /*return*/, results.Items];
                }
            });
        });
    };
    _RegionalSettings = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("regionalsettings")
    ], _RegionalSettings);
    return _RegionalSettings;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableInstance"]));

var RegionalSettings = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_RegionalSettings);
var _TimeZone = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_TimeZone, _super);
    function _TimeZone() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets an Local Time by UTC Time
     *
     * @param utcTime UTC Time as Date or ISO String
     */
    _TimeZone.prototype.utcToLocalTime = function (utcTime) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var dateIsoString, res;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (typeof utcTime === "string") {
                            dateIsoString = utcTime;
                        }
                        else {
                            dateIsoString = utcTime.toISOString();
                        }
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(TimeZone, "utctolocaltime('" + dateIsoString + "')"))];
                    case 1:
                        res = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(res, "UTCToLocalTime") ? res.UTCToLocalTime : res];
                }
            });
        });
    };
    /**
     * Gets an UTC Time by Local Time
     *
     * @param localTime Local Time as Date or ISO String
     */
    _TimeZone.prototype.localTimeToUTC = function (localTime) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var dateIsoString, res;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (typeof localTime === "string") {
                            dateIsoString = localTime;
                        }
                        else {
                            dateIsoString = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["dateAdd"])(localTime, "minute", localTime.getTimezoneOffset() * -1).toISOString();
                        }
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(TimeZone, "localtimetoutc('" + dateIsoString + "')"))];
                    case 1:
                        res = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(res, "LocalTimeToUTC") ? res.LocalTimeToUTC : res];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("tz.utcToLocalTime")
    ], _TimeZone.prototype, "utcToLocalTime", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("tz.localTimeToUTC")
    ], _TimeZone.prototype, "localTimeToUTC", null);
    _TimeZone = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("timezone")
    ], _TimeZone);
    return _TimeZone;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableInstance"]));

var TimeZone = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_TimeZone);
var _TimeZones = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_TimeZones, _super);
    function _TimeZones() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets an TimeZone by id (see: https://msdn.microsoft.com/en-us/library/office/jj247008.aspx)
     *
     * @param id The integer id of the timezone to retrieve
     */
    _TimeZones.prototype.getById = function (id) {
        // do the post and merge the result into a TimeZone instance so the data and methods are available
        return Object(_operations__WEBPACK_IMPORTED_MODULE_5__["spPost"])(this.clone(TimeZones, "GetById(" + id + ")").usingParser(Object(_odata__WEBPACK_IMPORTED_MODULE_4__["spODataEntity"])(TimeZone)));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("tzs.getById")
    ], _TimeZones.prototype, "getById", null);
    _TimeZones = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_3__["defaultPath"])("timezones")
    ], _TimeZones);
    return _TimeZones;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var TimeZones = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_TimeZones);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "lzYu":
/*!******************************************************************!*\
  !*** ./node_modules/@pnp/common/node_modules/tslib/tslib.es6.js ***!
  \******************************************************************/
/*! exports provided: __extends, __assign, __rest, __decorate, __param, __metadata, __awaiter, __generator, __exportStar, __values, __read, __spread, __spreadArrays, __await, __asyncGenerator, __asyncDelegator, __asyncValues, __makeTemplateObject, __importStar, __importDefault */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__extends", function() { return __extends; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__assign", function() { return __assign; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__rest", function() { return __rest; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__decorate", function() { return __decorate; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__param", function() { return __param; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__metadata", function() { return __metadata; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__awaiter", function() { return __awaiter; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__generator", function() { return __generator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__exportStar", function() { return __exportStar; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__values", function() { return __values; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__read", function() { return __read; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__spread", function() { return __spread; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__spreadArrays", function() { return __spreadArrays; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__await", function() { return __await; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncGenerator", function() { return __asyncGenerator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncDelegator", function() { return __asyncDelegator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncValues", function() { return __asyncValues; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__makeTemplateObject", function() { return __makeTemplateObject; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__importStar", function() { return __importStar; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__importDefault", function() { return __importDefault; });
/*! *****************************************************************************
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use
this file except in compliance with the License. You may obtain a copy of the
License at http://www.apache.org/licenses/LICENSE-2.0

THIS CODE IS PROVIDED ON AN *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, EITHER EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION ANY IMPLIED
WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE,
MERCHANTABLITY OR NON-INFRINGEMENT.

See the Apache Version 2.0 License for specific language governing permissions
and limitations under the License.
***************************************************************************** */
/* global Reflect, Promise */

var extendStatics = function(d, b) {
    extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return extendStatics(d, b);
};

function __extends(d, b) {
    extendStatics(d, b);
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
}

var __assign = function() {
    __assign = Object.assign || function __assign(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
        }
        return t;
    }
    return __assign.apply(this, arguments);
}

function __rest(s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
}

function __decorate(decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
}

function __param(paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
}

function __metadata(metadataKey, metadataValue) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(metadataKey, metadataValue);
}

function __awaiter(thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
}

function __generator(thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
}

function __exportStar(m, exports) {
    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
}

function __values(o) {
    var m = typeof Symbol === "function" && o[Symbol.iterator], i = 0;
    if (m) return m.call(o);
    return {
        next: function () {
            if (o && i >= o.length) o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
}

function __read(o, n) {
    var m = typeof Symbol === "function" && o[Symbol.iterator];
    if (!m) return o;
    var i = m.call(o), r, ar = [], e;
    try {
        while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
    }
    catch (error) { e = { error: error }; }
    finally {
        try {
            if (r && !r.done && (m = i["return"])) m.call(i);
        }
        finally { if (e) throw e.error; }
    }
    return ar;
}

function __spread() {
    for (var ar = [], i = 0; i < arguments.length; i++)
        ar = ar.concat(__read(arguments[i]));
    return ar;
}

function __spreadArrays() {
    for (var s = 0, i = 0, il = arguments.length; i < il; i++) s += arguments[i].length;
    for (var r = Array(s), k = 0, i = 0; i < il; i++)
        for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
            r[k] = a[j];
    return r;
};

function __await(v) {
    return this instanceof __await ? (this.v = v, this) : new __await(v);
}

function __asyncGenerator(thisArg, _arguments, generator) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var g = generator.apply(thisArg, _arguments || []), i, q = [];
    return i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i;
    function verb(n) { if (g[n]) i[n] = function (v) { return new Promise(function (a, b) { q.push([n, v, a, b]) > 1 || resume(n, v); }); }; }
    function resume(n, v) { try { step(g[n](v)); } catch (e) { settle(q[0][3], e); } }
    function step(r) { r.value instanceof __await ? Promise.resolve(r.value.v).then(fulfill, reject) : settle(q[0][2], r); }
    function fulfill(value) { resume("next", value); }
    function reject(value) { resume("throw", value); }
    function settle(f, v) { if (f(v), q.shift(), q.length) resume(q[0][0], q[0][1]); }
}

function __asyncDelegator(o) {
    var i, p;
    return i = {}, verb("next"), verb("throw", function (e) { throw e; }), verb("return"), i[Symbol.iterator] = function () { return this; }, i;
    function verb(n, f) { i[n] = o[n] ? function (v) { return (p = !p) ? { value: __await(o[n](v)), done: n === "return" } : f ? f(v) : v; } : f; }
}

function __asyncValues(o) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var m = o[Symbol.asyncIterator], i;
    return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
    function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
    function settle(resolve, reject, d, v) { Promise.resolve(v).then(function(v) { resolve({ value: v, done: d }); }, reject); }
}

function __makeTemplateObject(cooked, raw) {
    if (Object.defineProperty) { Object.defineProperty(cooked, "raw", { value: raw }); } else { cooked.raw = raw; }
    return cooked;
};

function __importStar(mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result.default = mod;
    return result;
}

function __importDefault(mod) {
    return (mod && mod.__esModule) ? mod : { default: mod };
}


/***/ }),

/***/ "pG6f":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/odata/invokable-binder.js ***!
  \*****************************************************/
/*! exports provided: invokableFactory */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "invokableFactory", function() { return invokableFactory; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "t3D1");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _invokable_extensions__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./invokable-extensions */ "01TU");



var invokableBinder = function (invoker) { return function (constructor) {
    return function () {
        var args = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            args[_i] = arguments[_i];
        }
        var factory = function (as) {
            var r = Object.assign(function () {
                var ags = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    ags[_i] = arguments[_i];
                }
                return invoker.apply(r, ags);
            }, new (constructor.bind.apply(constructor, Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__spread"])([void 0], as)))());
            Reflect.setPrototypeOf(r, constructor.prototype);
            return r;
        };
        if (_pnp_common__WEBPACK_IMPORTED_MODULE_1__["RuntimeConfig"].ie11) {
            return factory(args);
        }
        else {
            return new Proxy(Object(_invokable_extensions__WEBPACK_IMPORTED_MODULE_2__["applyFactoryExtensions"])(factory, args), {
                apply: function (target, _thisArg, argArray) {
                    return Object(_invokable_extensions__WEBPACK_IMPORTED_MODULE_2__["extensionOrDefault"])("apply", function () {
                        var a = [];
                        for (var _i = 0; _i < arguments.length; _i++) {
                            a[_i] = arguments[_i];
                        }
                        return Reflect.apply(a[0], a[1], a[2]);
                    }, target, _thisArg, argArray);
                },
                get: function (target, p, receiver) {
                    return Object(_invokable_extensions__WEBPACK_IMPORTED_MODULE_2__["extensionOrDefault"])("get", function () {
                        var a = [];
                        for (var _i = 0; _i < arguments.length; _i++) {
                            a[_i] = arguments[_i];
                        }
                        return Reflect.get(a[0], a[1], a[2]);
                    }, target, p, receiver);
                },
                has: function (target, p) {
                    return Object(_invokable_extensions__WEBPACK_IMPORTED_MODULE_2__["extensionOrDefault"])("has", function () {
                        var a = [];
                        for (var _i = 0; _i < arguments.length; _i++) {
                            a[_i] = arguments[_i];
                        }
                        return Reflect.has(a[0], a[1]);
                    }, target, p);
                },
                set: function (target, p, value, receiver) {
                    return Object(_invokable_extensions__WEBPACK_IMPORTED_MODULE_2__["extensionOrDefault"])("set", function () {
                        var a = [];
                        for (var _i = 0; _i < arguments.length; _i++) {
                            a[_i] = arguments[_i];
                        }
                        return Reflect.set(a[0], a[1], a[2], a[3]);
                    }, target, p, value, receiver);
                },
            });
        }
    };
}; };
var invokableFactory = invokableBinder(function (options) {
    return this.defaultAction(options);
});
//# sourceMappingURL=invokable-binder.js.map

/***/ }),

/***/ "pGaT":
/*!*******************************************************!*\
  !*** ./node_modules/@pnp/sp/column-defaults/index.js ***!
  \*******************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./list */ "jM62");
/* harmony import */ var _folder__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./folder */ "P6NM");


//# sourceMappingURL=index.js.map

/***/ }),

/***/ "pGs8":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/files/item.js ***!
  \********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _items_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../items/types */ "3DT9");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "qfcF");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_items_types__WEBPACK_IMPORTED_MODULE_1__["_Item"], "file", _types__WEBPACK_IMPORTED_MODULE_2__["File"], "file");
//# sourceMappingURL=item.js.map

/***/ }),

/***/ "pM8y":
/*!******************************************************************!*\
  !*** ./lib/extensions/announcements/components/Announcements.js ***!
  \******************************************************************/
/*! exports provided: default */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "default", function() { return RenderAnnouncements; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "cDcd");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _microsoft_office_ui_fabric_react_bundle__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @microsoft/office-ui-fabric-react-bundle */ "KL1q");
/* harmony import */ var _microsoft_office_ui_fabric_react_bundle__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(_microsoft_office_ui_fabric_react_bundle__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _pnp_sp_presets_all__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/sp/presets/all */ "F/m7");
/* harmony import */ var announcementsStrings__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! announcementsStrings */ "+n6P");
/* harmony import */ var announcementsStrings__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(announcementsStrings__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var _AnnouncementsApplicationCustomizer__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../AnnouncementsApplicationCustomizer */ "wdEC");






function RenderAnnouncements(props) {
    // Two local state variables with their setter
    var _a = Object(react__WEBPACK_IMPORTED_MODULE_0__["useState"])([]), announcements = _a[0], setAnnouncements = _a[1];
    var _b = Object(react__WEBPACK_IMPORTED_MODULE_0__["useState"])([]), acknowledgedAnnouncements = _b[0], setAcknowledgedAnnouncements = _b[1];
    // Use an effect to query the list data only once,
    // not on every render. The effect will be re-run if
    // props.siteUrl or props.listName changes
    Object(react__WEBPACK_IMPORTED_MODULE_0__["useEffect"])(function () {
        if (window.localStorage) {
            var items = window.localStorage.getItem(_AnnouncementsApplicationCustomizer__WEBPACK_IMPORTED_MODULE_4__["QUALIFIED_NAME"]);
            if (items) {
                setAcknowledgedAnnouncements(JSON.parse(items));
            }
        }
        // Use PnP JS to query SharePoint
        var now = new Date().toISOString();
        Object(_pnp_sp_presets_all__WEBPACK_IMPORTED_MODULE_2__["Web"])(props.siteUrl)
            .lists.getByTitle(props.listName)
            .items
            .filter("(Locale eq '" + props.culture + "' or Locale eq null) and (StartDateTime le datetime'" + now + "' or StartDateTime eq null) and (EndDateTime ge datetime'" + now + "' or EndDateTime eq null)")
            .select("ID", "Title", "Announcement", "Urgent", "Link", "Locale", "StartDateTime", "EndDateTime")
            .get()
            .then(setAnnouncements);
    }, [props.siteUrl, props.listName]);
    var announcementElements = announcements
        .filter(function (announcement) { return acknowledgedAnnouncements.indexOf(announcement.ID) < 0; })
        .map(function (announcement) { return react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_microsoft_office_ui_fabric_react_bundle__WEBPACK_IMPORTED_MODULE_1__["MessageBar"], { messageBarType: (announcement.Urgent ? _microsoft_office_ui_fabric_react_bundle__WEBPACK_IMPORTED_MODULE_1__["MessageBarType"].error : _microsoft_office_ui_fabric_react_bundle__WEBPACK_IMPORTED_MODULE_1__["MessageBarType"].warning), isMultiline: false, onDismiss: function () {
            // On dismiss, add the current announcement id to the array 
            // STORAGE_KEY item in localStorage so it is remembered locally
            var items = JSON.parse(window.localStorage.getItem(_AnnouncementsApplicationCustomizer__WEBPACK_IMPORTED_MODULE_4__["QUALIFIED_NAME"])) || [];
            items.push(announcement.ID);
            window.localStorage.setItem(_AnnouncementsApplicationCustomizer__WEBPACK_IMPORTED_MODULE_4__["QUALIFIED_NAME"], JSON.stringify(items));
            setAcknowledgedAnnouncements(items);
        }, dismissButtonAriaLabel: announcementsStrings__WEBPACK_IMPORTED_MODULE_3__["Close"] },
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("strong", null, announcement.Title),
        "\u00A0",
        react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("span", { dangerouslySetInnerHTML: { __html: announcement.Announcement } }),
        announcement.Link && react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_microsoft_office_ui_fabric_react_bundle__WEBPACK_IMPORTED_MODULE_1__["Link"], { href: announcement.Link.Url, target: "_blank" }, announcement.Link.Description)); });
    return react__WEBPACK_IMPORTED_MODULE_0__["createElement"]("div", null, announcementElements);
}


/***/ }),

/***/ "pNrY":
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/regional-settings/index.js ***!
  \*********************************************************/
/*! exports provided: RegionalSettings, TimeZone, TimeZones */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "JT9m");
/* harmony import */ var _user_custom_actions__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./user-custom-actions */ "2B8P");
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./list */ "KJ7d");
/* harmony import */ var _field__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./field */ "DL6e");
/* harmony import */ var _content_type__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./content-type */ "auja");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./types */ "luMW");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RegionalSettings", function() { return _types__WEBPACK_IMPORTED_MODULE_5__["RegionalSettings"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "TimeZone", function() { return _types__WEBPACK_IMPORTED_MODULE_5__["TimeZone"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "TimeZones", function() { return _types__WEBPACK_IMPORTED_MODULE_5__["TimeZones"]; });







//# sourceMappingURL=index.js.map

/***/ }),

/***/ "pTlN":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/navigation/web.js ***!
  \************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "BK4w");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "navigation", _types__WEBPACK_IMPORTED_MODULE_2__["Navigation"]);
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "pXI8":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/webparts/types.js ***!
  \************************************************/
/*! exports provided: _LimitedWebPartManager, LimitedWebPartManager, _WebPartDefinitions, WebPartDefinitions, _WebPartDefinition, WebPartDefinition, WebPartsPersonalizationScope */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_LimitedWebPartManager", function() { return _LimitedWebPartManager; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "LimitedWebPartManager", function() { return LimitedWebPartManager; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_WebPartDefinitions", function() { return _WebPartDefinitions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "WebPartDefinitions", function() { return WebPartDefinitions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_WebPartDefinition", function() { return _WebPartDefinition; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "WebPartDefinition", function() { return WebPartDefinition; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "WebPartsPersonalizationScope", function() { return WebPartsPersonalizationScope; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../telemetry */ "h3gW");





var _LimitedWebPartManager = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_LimitedWebPartManager, _super);
    function _LimitedWebPartManager() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(_LimitedWebPartManager.prototype, "scope", {
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_4__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryable"])(this, "Scope"), "f.scope");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_LimitedWebPartManager.prototype, "webparts", {
        get: function () {
            return WebPartDefinitions(this, "webparts");
        },
        enumerable: true,
        configurable: true
    });
    _LimitedWebPartManager.prototype.export = function (id) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(LimitedWebPartManagerCloneFactory, "ExportWebPart"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({ webPartId: id }));
    };
    _LimitedWebPartManager.prototype.import = function (xml) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(LimitedWebPartManagerCloneFactory, "ImportWebPart"), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({ webPartXml: xml }));
    };
    return _LimitedWebPartManager;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryable"]));

var LimitedWebPartManager = function (baseUrl, path) { return new _LimitedWebPartManager(baseUrl, path); };
var LimitedWebPartManagerCloneFactory = function (baseUrl, path) { return LimitedWebPartManager(baseUrl, path); };
var _WebPartDefinitions = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_WebPartDefinitions, _super);
    function _WebPartDefinitions() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a web part definition from the collection by id
     *
     * @param id The storage ID of the SPWebPartDefinition to retrieve
     */
    _WebPartDefinitions.prototype.getById = function (id) {
        return WebPartDefinition(this, "getbyid('" + id + "')");
    };
    /**
     * Gets a web part definition from the collection by storage id
     *
     * @param id The WebPart.ID of the SPWebPartDefinition to retrieve
     */
    _WebPartDefinitions.prototype.getByControlId = function (id) {
        return WebPartDefinition(this, "getByControlId('" + id + "')");
    };
    return _WebPartDefinitions;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var WebPartDefinitions = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_WebPartDefinitions);
var _WebPartDefinition = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_WebPartDefinition, _super);
    function _WebPartDefinition() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(_WebPartDefinition.prototype, "webpart", {
        /**
        * Gets the webpart information associated with this definition
        */
        get: function () {
            return Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryableInstance"])(this, "webpart");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Saves changes to the Web Part made using other properties and methods on the SPWebPartDefinition object
     */
    _WebPartDefinition.prototype.saveChanges = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(WebPartDefinition, "SaveWebPartChanges"));
    };
    /**
     * Moves the Web Part to a different location on a Web Part Page
     *
     * @param zoneId The ID of the Web Part Zone to which to move the Web Part
     * @param zoneIndex A Web Part zone index that specifies the position at which the Web Part is to be moved within the destination Web Part zone
     */
    _WebPartDefinition.prototype.moveTo = function (zoneId, zoneIndex) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(WebPartDefinition, "MoveWebPartTo(zoneID='" + zoneId + "', zoneIndex=" + zoneIndex + ")"));
    };
    /**
     * Closes the Web Part. If the Web Part is already closed, this method does nothing
     */
    _WebPartDefinition.prototype.close = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(WebPartDefinition, "CloseWebPart"));
    };
    /**
     * Opens the Web Part. If the Web Part is already closed, this method does nothing
     */
    _WebPartDefinition.prototype.open = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(WebPartDefinition, "OpenWebPart"));
    };
    /**
     * Removes a webpart from a page, all settings will be lost
     */
    _WebPartDefinition.prototype.delete = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_3__["spPost"])(this.clone(WebPartDefinition, "DeleteWebPart"));
    };
    return _WebPartDefinition;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var WebPartDefinition = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_WebPartDefinition);
var WebPartsPersonalizationScope;
(function (WebPartsPersonalizationScope) {
    WebPartsPersonalizationScope[WebPartsPersonalizationScope["User"] = 0] = "User";
    WebPartsPersonalizationScope[WebPartsPersonalizationScope["Shared"] = 1] = "Shared";
})(WebPartsPersonalizationScope || (WebPartsPersonalizationScope = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "psVa":
/*!********************************************!*\
  !*** ./node_modules/@pnp/sp/fields/web.js ***!
  \********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "KYlJ");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "fields", _types__WEBPACK_IMPORTED_MODULE_2__["Fields"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "availablefields", _types__WEBPACK_IMPORTED_MODULE_2__["Fields"], "availablefields");
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "q1Lp":
/*!***********************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/escapeQueryStrValue.js ***!
  \***********************************************************/
/*! exports provided: escapeQueryStrValue */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "escapeQueryStrValue", function() { return escapeQueryStrValue; });
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _pnp_logging__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/logging */ "B3L+");


function escapeQueryStrValue(value) {
    if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["stringIsNullOrEmpty"])(value)) {
        return "";
    }
    // replace all instance of ' with ''
    if (/!(@.*?)::(.*?)/ig.test(value)) {
        // to ensure our param aliasing still works we need to treat these special or we'll hear about it
        // so we encode JUST the part that will end up in the url
        return value.replace(/!(@.*?)::(.*)$/ig, function (match, labelName, v) {
            _pnp_logging__WEBPACK_IMPORTED_MODULE_1__["Logger"].write("Rewriting aliased parameter from match " + match + " to label: " + labelName + " value: " + v, 0 /* Verbose */);
            return "!" + labelName + "::" + encodeURIComponent(v.replace(/\'/ig, "''"));
        });
    }
    else {
        return encodeURIComponent(value.replace(/\'/ig, "''"));
    }
}
//# sourceMappingURL=escapeQueryStrValue.js.map

/***/ }),

/***/ "qPin":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/sharing/file.js ***!
  \**********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _files_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../files/types */ "qfcF");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "AvPO");
/* harmony import */ var _funcs__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./funcs */ "At/s");



_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.shareWith = function (loginNames, role, requireSignin, emailData) {
    if (role === void 0) { role = _types__WEBPACK_IMPORTED_MODULE_1__["SharingRole"].View; }
    if (requireSignin === void 0) { requireSignin = false; }
    return Object(_funcs__WEBPACK_IMPORTED_MODULE_2__["shareWith"])(this, loginNames, role, requireSignin, false, emailData);
};
_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.getShareLink = _funcs__WEBPACK_IMPORTED_MODULE_2__["getShareLink"];
_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.checkSharingPermissions = _funcs__WEBPACK_IMPORTED_MODULE_2__["checkPermissions"];
_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.getSharingInformation = _funcs__WEBPACK_IMPORTED_MODULE_2__["getSharingInformation"];
_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.getObjectSharingSettings = _funcs__WEBPACK_IMPORTED_MODULE_2__["getObjectSharingSettings"];
_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.unshare = _funcs__WEBPACK_IMPORTED_MODULE_2__["unshareObject"];
_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.deleteSharingLinkByKind = _funcs__WEBPACK_IMPORTED_MODULE_2__["deleteLinkByKind"];
_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.unshareLink = _funcs__WEBPACK_IMPORTED_MODULE_2__["unshareLink"];
//# sourceMappingURL=file.js.map

/***/ }),

/***/ "qfcF":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/files/types.js ***!
  \*********************************************/
/*! exports provided: _Files, Files, _File, File, _Versions, Versions, _Version, Version, CheckinType, MoveOperations, TemplateFileType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Files", function() { return _Files; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Files", function() { return Files; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_File", function() { return _File; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "File", function() { return File; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Versions", function() { return _Versions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Versions", function() { return Versions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Version", function() { return _Version; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Version", function() { return Version; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "CheckinType", function() { return CheckinType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "MoveOperations", function() { return MoveOperations; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TemplateFileType", function() { return TemplateFileType; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _items__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../items */ "lYrR");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../telemetry */ "h3gW");
/* harmony import */ var _utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! ../utils/toResourcePath */ "t8hP");












/**
 * Describes a collection of File objects
 *
 */
var _Files = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Files, _super);
    function _Files() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a File by filename
     *
     * @param name The name of the file, including extension.
     */
    _Files.prototype.getByName = function (name) {
        if (/\%#/.test(name)) {
            throw Error("For file names containing % or # please use web.getFileByServerRelativePath");
        }
        return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(File(this).concat("('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(name) + "')"), "fis.getByName");
    };
    /**
     * Uploads a file. Not supported for batching
     *
     * @param url The folder-relative url of the file.
     * @param content The file contents blob.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
     * @returns The new File and the raw response.
     */
    _Files.prototype.add = function (url, content, shouldOverWrite) {
        if (shouldOverWrite === void 0) { shouldOverWrite = true; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var response;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Files(this, "add(overwrite=" + shouldOverWrite + ",url='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(url) + "')"), {
                            body: content,
                        })];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, {
                                data: response,
                                file: this.getByName(url),
                            }];
                }
            });
        });
    };
    /**
     * Adds a file using the pound percent safe methods
     *
     * @param url Excoded url of the file
     * @param content The file content
     * @param parameters Additional parameters to control method behavior
     */
    _Files.prototype.addUsingPath = function (url, content, parameters) {
        if (parameters === void 0) { parameters = { Overwrite: false }; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var path, resp;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        path = ["AddUsingPath(decodedurl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(url) + "'"];
                        if (parameters) {
                            if (parameters.Overwrite) {
                                path.push(",Overwrite=true");
                            }
                            if (parameters.AutoCheckoutOnInvalidData) {
                                path.push(",AutoCheckoutOnInvalidData=true");
                            }
                            if (!Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["stringIsNullOrEmpty"])(parameters.XorHash)) {
                                path.push(",XorHash=" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(parameters.XorHash));
                            }
                        }
                        path.push(")");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Files(this, path.join("")), { body: content })];
                    case 1:
                        resp = _a.sent();
                        return [2 /*return*/, {
                                data: resp,
                                file: File(Object(_odata__WEBPACK_IMPORTED_MODULE_5__["odataUrlFrom"])(resp)),
                            }];
                }
            });
        });
    };
    /**
     * Uploads a file. Not supported for batching
     *
     * @param url The folder-relative url of the file.
     * @param content The Blob file content to add
     * @param progress A callback function which can be used to track the progress of the upload
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten? (default: true)
     * @param chunkSize The size of each file slice, in bytes (default: 10485760)
     * @returns The new File and the raw response.
     */
    _Files.prototype.addChunked = function (url, content, progress, shouldOverWrite, chunkSize) {
        if (shouldOverWrite === void 0) { shouldOverWrite = true; }
        if (chunkSize === void 0) { chunkSize = 10485760; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var file;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Files, "add(overwrite=" + shouldOverWrite + ",url='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(url) + "')", false))];
                    case 1:
                        _a.sent();
                        file = this.getByName(url);
                        return [4 /*yield*/, file.setContentChunked(content, progress, chunkSize)];
                    case 2: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    /**
     * Adds a ghosted file to an existing list or document library. Not supported for batching.
     *
     * @param fileUrl The server-relative url where you want to save the file.
     * @param templateFileType The type of use to create the file.
     * @returns The template file that was added and the raw response.
     */
    _Files.prototype.addTemplateFile = function (fileUrl, templateFileType) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var response;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Files, "addTemplateFile(urloffile='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(fileUrl) + "',templatefiletype=" + templateFileType + ")", false))];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, {
                                data: response,
                                file: File(Object(_odata__WEBPACK_IMPORTED_MODULE_5__["odataUrlFrom"])(response)),
                            }];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fis.add")
    ], _Files.prototype, "add", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fis.addUsingPath")
    ], _Files.prototype, "addUsingPath", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fis.addChunked")
    ], _Files.prototype, "addChunked", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fis.addTemplateFile")
    ], _Files.prototype, "addTemplateFile", null);
    _Files = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("files")
    ], _Files);
    return _Files;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var Files = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Files);
/**
 * Describes a single File instance
 *
 */
var _File = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_File, _super);
    function _File() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["deleteableWithETag"])("fi");
        return _this;
    }
    Object.defineProperty(_File.prototype, "listItemAllFields", {
        /**
         * Gets a value that specifies the list item field values for the list item corresponding to the file.
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["SharePointQueryableInstance"])(this, "listItemAllFields"), "fi.listItemAllFields");
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(_File.prototype, "versions", {
        /**
         * Gets a collection of versions
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Versions(this), "fi.versions");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Approves the file submitted for content approval with the specified comment.
     * Only documents in lists that are enabled for content approval can be approved.
     *
     * @param comment The comment for the approval.
     */
    _File.prototype.approve = function (comment) {
        if (comment === void 0) { comment = ""; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "approve(comment='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(comment) + "')"));
    };
    /**
     * Stops the chunk upload session without saving the uploaded data. Does not support batching.
     * If the file doesn’t already exist in the library, the partially uploaded file will be deleted.
     * Use this in response to user action (as in a request to cancel an upload) or an error or exception.
     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     */
    _File.prototype.cancelUpload = function (uploadId) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "cancelUpload(uploadId=guid'" + uploadId + "')", false));
    };
    /**
     * Checks the file in to a document library based on the check-in type.
     *
     * @param comment A comment for the check-in. Its length must be <= 1023.
     * @param checkinType The check-in type for the file.
     */
    _File.prototype.checkin = function (comment, checkinType) {
        if (comment === void 0) { comment = ""; }
        if (checkinType === void 0) { checkinType = CheckinType.Major; }
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "checkin(comment='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(comment) + "',checkintype=" + checkinType + ")"));
    };
    /**
     * Checks out the file from a document library.
     */
    _File.prototype.checkout = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "checkout"));
    };
    /**
     * Copies the file to the destination url.
     *
     * @param url The absolute url or server relative url of the destination file path to copy to.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
     */
    _File.prototype.copyTo = function (url, shouldOverWrite) {
        if (shouldOverWrite === void 0) { shouldOverWrite = true; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "copyTo(strnewurl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(url) + "',boverwrite=" + shouldOverWrite + ")"));
    };
    /**
     * Copies the file by path to destination path.
     * Also works with different site collections.
     *
     * @param destUrl The absolute url or server relative url of the destination file path to copy to.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
     * @param keepBoth Keep both if file with the same name in the same location already exists? Only relevant when shouldOverWrite is set to false.
     */
    _File.prototype.copyByPath = function (destUrl, shouldOverWrite, KeepBoth) {
        if (KeepBoth === void 0) { KeepBoth = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var _a, srcUrl, absoluteUrl, webBaseUrl, hostUrl;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.select("ServerRelativeUrl")()];
                    case 1:
                        _a = _b.sent(), srcUrl = _a.ServerRelativeUrl, absoluteUrl = _a["odata.id"];
                        webBaseUrl = Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(absoluteUrl);
                        hostUrl = webBaseUrl.replace("://", "___").split("/")[0].replace("___", "://");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(File(webBaseUrl, "/_api/SP.MoveCopyUtil.CopyFileByPath(overwrite=@a1)?@a1=" + shouldOverWrite), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({
                                destPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__["toResourcePath"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["isUrlAbsolute"])(destUrl) ? destUrl : "" + hostUrl + destUrl),
                                options: {
                                    KeepBoth: KeepBoth,
                                    ResetAuthorAndCreatedOnCopy: true,
                                    ShouldBypassSharedLocks: true,
                                    __metadata: {
                                        type: "SP.MoveCopyOptions",
                                    },
                                },
                                srcPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__["toResourcePath"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["isUrlAbsolute"])(srcUrl) ? srcUrl : "" + hostUrl + srcUrl),
                            }))];
                    case 2:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Denies approval for a file that was submitted for content approval.
     * Only documents in lists that are enabled for content approval can be denied.
     *
     * @param comment The comment for the denial.
     */
    _File.prototype.deny = function (comment) {
        if (comment === void 0) { comment = ""; }
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "deny(comment='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(comment) + "')"));
    };
    /**
     * Moves the file to the specified destination url.
     *
     * @param url The absolute url or server relative url of the destination file path to move to.
     * @param moveOperations The bitwise MoveOperations value for how to move the file.
     */
    _File.prototype.moveTo = function (url, moveOperations) {
        if (moveOperations === void 0) { moveOperations = MoveOperations.Overwrite; }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "moveTo(newurl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(url) + "',flags=" + moveOperations + ")"));
    };
    /**
     * Moves the file by path to the specified destination url.
     * Also works with different site collections.
     *
     * @param destUrl The absolute url or server relative url of the destination file path to move to.
     * @param shouldOverWrite Should a file with the same name in the same location be overwritten?
     * @param keepBoth Keep both if file with the same name in the same location already exists? Only relevant when shouldOverWrite is set to false.
     */
    _File.prototype.moveByPath = function (destUrl, shouldOverWrite, KeepBoth) {
        if (KeepBoth === void 0) { KeepBoth = false; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var _a, srcUrl, absoluteUrl, webBaseUrl, hostUrl;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_b) {
                switch (_b.label) {
                    case 0: return [4 /*yield*/, this.select("ServerRelativeUrl")()];
                    case 1:
                        _a = _b.sent(), srcUrl = _a.ServerRelativeUrl, absoluteUrl = _a["odata.id"];
                        webBaseUrl = Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_9__["extractWebUrl"])(absoluteUrl);
                        hostUrl = webBaseUrl.replace("://", "___").split("/")[0].replace("___", "://");
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(File(webBaseUrl, "/_api/SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=" + shouldOverWrite), Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["body"])({
                                destPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__["toResourcePath"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["isUrlAbsolute"])(destUrl) ? destUrl : "" + hostUrl + destUrl),
                                options: {
                                    KeepBoth: KeepBoth,
                                    ResetAuthorAndCreatedOnCopy: false,
                                    ShouldBypassSharedLocks: true,
                                    __metadata: {
                                        type: "SP.MoveCopyOptions",
                                    },
                                },
                                srcPath: Object(_utils_toResourcePath__WEBPACK_IMPORTED_MODULE_11__["toResourcePath"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["isUrlAbsolute"])(srcUrl) ? srcUrl : "" + hostUrl + srcUrl),
                            }))];
                    case 2:
                        _b.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Submits the file for content approval with the specified comment.
     *
     * @param comment The comment for the published file. Its length must be <= 1023.
     */
    _File.prototype.publish = function (comment) {
        if (comment === void 0) { comment = ""; }
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "publish(comment='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(comment) + "')"));
    };
    /**
     * Moves the file to the Recycle Bin and returns the identifier of the new Recycle Bin item.
     *
     * @returns The GUID of the recycled file.
     */
    _File.prototype.recycle = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "recycle"));
    };
    /**
     * Reverts an existing checkout for the file.
     *
     */
    _File.prototype.undoCheckout = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "undoCheckout"));
    };
    /**
     * Removes the file from content approval or unpublish a major version.
     *
     * @param comment The comment for the unpublish operation. Its length must be <= 1023.
     */
    _File.prototype.unpublish = function (comment) {
        if (comment === void 0) { comment = ""; }
        if (comment.length > 1023) {
            throw Error("The maximum comment length is 1023 characters.");
        }
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "unpublish(comment='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(comment) + "')"));
    };
    /**
     * Gets the contents of the file as text. Not supported in batching.
     *
     */
    _File.prototype.getText = function () {
        return this.clone(File, "$value", false).usingParser(new _pnp_odata__WEBPACK_IMPORTED_MODULE_2__["TextParser"]())(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "binaryStringResponseBody": "true" }));
    };
    /**
     * Gets the contents of the file as a blob, does not work in Node.js. Not supported in batching.
     *
     */
    _File.prototype.getBlob = function () {
        return this.clone(File, "$value", false).usingParser(new _pnp_odata__WEBPACK_IMPORTED_MODULE_2__["BlobParser"]())(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "binaryStringResponseBody": "true" }));
    };
    /**
     * Gets the contents of a file as an ArrayBuffer, works in Node.js. Not supported in batching.
     */
    _File.prototype.getBuffer = function () {
        return this.clone(File, "$value", false).usingParser(new _pnp_odata__WEBPACK_IMPORTED_MODULE_2__["BufferParser"]())(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "binaryStringResponseBody": "true" }));
    };
    /**
     * Gets the contents of a file as an ArrayBuffer, works in Node.js. Not supported in batching.
     */
    _File.prototype.getJSON = function () {
        return this.clone(File, "$value", false).usingParser(new _pnp_odata__WEBPACK_IMPORTED_MODULE_2__["JSONParser"]())(Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_2__["headers"])({ "binaryStringResponseBody": "true" }));
    };
    /**
     * Sets the content of a file, for large files use setContentChunked. Not supported in batching.
     *
     * @param content The file content
     *
     */
    _File.prototype.setContent = function (content) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "$value", false), {
                            body: content,
                            headers: {
                                "X-HTTP-Method": "PUT",
                            },
                        })];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, File(this)];
                }
            });
        });
    };
    /**
     * Gets the associated list item for this folder, loading the default properties
     */
    _File.prototype.getItem = function () {
        var selects = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            selects[_i] = arguments[_i];
        }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var q, d;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        q = this.listItemAllFields;
                        return [4 /*yield*/, q.select.apply(q, selects)()];
                    case 1:
                        d = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["assign"])(Object(_items__WEBPACK_IMPORTED_MODULE_4__["Item"])(Object(_odata__WEBPACK_IMPORTED_MODULE_5__["odataUrlFrom"])(d)), d)];
                }
            });
        });
    };
    /**
     * Sets the contents of a file using a chunked upload approach. Not supported in batching.
     *
     * @param file The file to upload
     * @param progress A callback function which can be used to track the progress of the upload
     * @param chunkSize The size of each file slice, in bytes (default: 10485760)
     */
    _File.prototype.setContentChunked = function (file, progress, chunkSize) {
        if (chunkSize === void 0) { chunkSize = 10485760; }
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var fileSize, totalBlocks, uploadId, currentPointer, i;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["isFunc"])(progress)) {
                            progress = function () { return null; };
                        }
                        fileSize = file.size;
                        totalBlocks = parseInt((fileSize / chunkSize).toString(), 10) + ((fileSize % chunkSize === 0) ? 1 : 0);
                        uploadId = Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["getGUID"])();
                        // report that we are starting
                        progress({ uploadId: uploadId, blockNumber: 1, chunkSize: chunkSize, currentPointer: 0, fileSize: fileSize, stage: "starting", totalBlocks: totalBlocks });
                        return [4 /*yield*/, this.startUpload(uploadId, file.slice(0, chunkSize))];
                    case 1:
                        currentPointer = _a.sent();
                        i = 2;
                        _a.label = 2;
                    case 2:
                        if (!(i < totalBlocks)) return [3 /*break*/, 5];
                        progress({ uploadId: uploadId, blockNumber: i, chunkSize: chunkSize, currentPointer: currentPointer, fileSize: fileSize, stage: "continue", totalBlocks: totalBlocks });
                        return [4 /*yield*/, this.continueUpload(uploadId, currentPointer, file.slice(currentPointer, currentPointer + chunkSize))];
                    case 3:
                        currentPointer = _a.sent();
                        _a.label = 4;
                    case 4:
                        i++;
                        return [3 /*break*/, 2];
                    case 5:
                        progress({ uploadId: uploadId, blockNumber: totalBlocks, chunkSize: chunkSize, currentPointer: currentPointer, fileSize: fileSize, stage: "finishing", totalBlocks: totalBlocks });
                        return [2 /*return*/, this.finishUpload(uploadId, currentPointer, file.slice(currentPointer))];
                }
            });
        });
    };
    /**
     * Starts a new chunk upload session and uploads the first fragment.
     * The current file content is not changed when this method completes.
     * The method is idempotent (and therefore does not change the result) as long as you use the same values for uploadId and stream.
     * The upload session ends either when you use the CancelUpload method or when you successfully
     * complete the upload session by passing the rest of the file contents through the ContinueUpload and FinishUpload methods.
     * The StartUpload and ContinueUpload methods return the size of the running total of uploaded data in bytes,
     * so you can pass those return values to subsequent uses of ContinueUpload and FinishUpload.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     * @param fragment The file contents.
     * @returns The size of the total uploaded data in bytes.
     */
    _File.prototype.startUpload = function (uploadId, fragment) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var n;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "startUpload(uploadId=guid'" + uploadId + "')", false), { body: fragment })];
                    case 1:
                        n = _a.sent();
                        if (typeof n === "object") {
                            // When OData=verbose the payload has the following shape:
                            // { StartUpload: "10485760" }
                            n = n.StartUpload;
                        }
                        return [2 /*return*/, parseFloat(n)];
                }
            });
        });
    };
    /**
     * Continues the chunk upload session with an additional fragment.
     * The current file content is not changed.
     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     * @param fileOffset The size of the offset into the file where the fragment starts.
     * @param fragment The file contents.
     * @returns The size of the total uploaded data in bytes.
     */
    _File.prototype.continueUpload = function (uploadId, fileOffset, fragment) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var n;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "continueUpload(uploadId=guid'" + uploadId + "',fileOffset=" + fileOffset + ")", false), { body: fragment })];
                    case 1:
                        n = _a.sent();
                        if (typeof n === "object") {
                            // When OData=verbose the payload has the following shape:
                            // { ContinueUpload: "20971520" }
                            n = n.ContinueUpload;
                        }
                        return [2 /*return*/, parseFloat(n)];
                }
            });
        });
    };
    /**
     * Uploads the last file fragment and commits the file. The current file content is changed when this method completes.
     * Use the uploadId value that was passed to the StartUpload method that started the upload session.
     * This method is currently available only on Office 365.
     *
     * @param uploadId The unique identifier of the upload session.
     * @param fileOffset The size of the offset into the file where the fragment starts.
     * @param fragment The file contents.
     * @returns The newly uploaded file.
     */
    _File.prototype.finishUpload = function (uploadId, fileOffset, fragment) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var response;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(File, "finishUpload(uploadId=guid'" + uploadId + "',fileOffset=" + fileOffset + ")", false), { body: fragment })];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, {
                                data: response,
                                file: File(Object(_odata__WEBPACK_IMPORTED_MODULE_5__["odataUrlFrom"])(response)),
                            }];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.approve")
    ], _File.prototype, "approve", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.cancelUpload")
    ], _File.prototype, "cancelUpload", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.checkin")
    ], _File.prototype, "checkin", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.checkout")
    ], _File.prototype, "checkout", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.copyTo")
    ], _File.prototype, "copyTo", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.copyByPath")
    ], _File.prototype, "copyByPath", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.deny")
    ], _File.prototype, "deny", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.moveTo")
    ], _File.prototype, "moveTo", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.moveByPath")
    ], _File.prototype, "moveByPath", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.publish")
    ], _File.prototype, "publish", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.recycle")
    ], _File.prototype, "recycle", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.undoCheckout")
    ], _File.prototype, "undoCheckout", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.unpublish")
    ], _File.prototype, "unpublish", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.getText")
    ], _File.prototype, "getText", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.getBlob")
    ], _File.prototype, "getBlob", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.getBuffer")
    ], _File.prototype, "getBuffer", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.getJSON")
    ], _File.prototype, "getJSON", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.setContent")
    ], _File.prototype, "setContent", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.getItem")
    ], _File.prototype, "getItem", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.startUpload")
    ], _File.prototype, "startUpload", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.continueUpload")
    ], _File.prototype, "continueUpload", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("fi.finishUpload")
    ], _File.prototype, "finishUpload", null);
    return _File;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var File = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_File);
/**
 * Describes a collection of Version objects
 *
 */
var _Versions = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Versions, _super);
    function _Versions() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a version by id
     *
     * @param versionId The id of the version to retrieve
     */
    _Versions.prototype.getById = function (versionId) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"].configure(Version(this).concat("(" + versionId + ")"), "vers.getById");
    };
    /**
     * Deletes all the file version objects in the collection.
     *
     */
    _Versions.prototype.deleteAll = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(Versions(this, "deleteAll"));
    };
    /**
     * Deletes the specified version of the file.
     *
     * @param versionId The ID of the file version to delete.
     */
    _Versions.prototype.deleteById = function (versionId) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Versions, "deleteById(vid=" + versionId + ")"));
    };
    /**
     * Recycles the specified version of the file.
     *
     * @param versionId The ID of the file version to delete.
     */
    _Versions.prototype.recycleByID = function (versionId) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Versions, "recycleByID(vid=" + versionId + ")"));
    };
    /**
     * Deletes the file version object with the specified version label.
     *
     * @param label The version label of the file version to delete, for example: 1.2
     */
    _Versions.prototype.deleteByLabel = function (label) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Versions, "deleteByLabel(versionlabel='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(label) + "')"));
    };
    /**
     * Recycles the file version object with the specified version label.
     *
     * @param label The version label of the file version to delete, for example: 1.2
     */
    _Versions.prototype.recycleByLabel = function (label) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Versions, "recycleByLabel(versionlabel='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(label) + "')"));
    };
    /**
     * Creates a new file version from the file specified by the version label.
     *
     * @param label The version label of the file version to restore, for example: 1.2
     */
    _Versions.prototype.restoreByLabel = function (label) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Versions, "restoreByLabel(versionlabel='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_8__["escapeQueryStrValue"])(label) + "')"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("vers.deleteAll")
    ], _Versions.prototype, "deleteAll", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("vers.deleteById")
    ], _Versions.prototype, "deleteById", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("vers.recycleByID")
    ], _Versions.prototype, "recycleByID", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("vers.deleteByLabel")
    ], _Versions.prototype, "deleteByLabel", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("vers.recycleByLabel")
    ], _Versions.prototype, "recycleByLabel", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_10__["tag"])("vers.restoreByLabel")
    ], _Versions.prototype, "restoreByLabel", null);
    _Versions = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("versions")
    ], _Versions);
    return _Versions;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var Versions = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Versions);
/**
 * Describes a single Version instance
 *
 */
var _Version = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Version, _super);
    function _Version() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["deleteableWithETag"])("ver");
        return _this;
    }
    return _Version;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Version = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Version);
/**
 * Types for document check in.
 * Minor = 0
 * Major = 1
 * Overwrite = 2
 */
var CheckinType;
(function (CheckinType) {
    CheckinType[CheckinType["Minor"] = 0] = "Minor";
    CheckinType[CheckinType["Major"] = 1] = "Major";
    CheckinType[CheckinType["Overwrite"] = 2] = "Overwrite";
})(CheckinType || (CheckinType = {}));
/**
 * File move opertions
 * Overwrite = 1
 * AllowBrokenThickets = 8
 */
var MoveOperations;
(function (MoveOperations) {
    MoveOperations[MoveOperations["Overwrite"] = 1] = "Overwrite";
    MoveOperations[MoveOperations["AllowBrokenThickets"] = 8] = "AllowBrokenThickets";
})(MoveOperations || (MoveOperations = {}));
var TemplateFileType;
(function (TemplateFileType) {
    TemplateFileType[TemplateFileType["StandardPage"] = 0] = "StandardPage";
    TemplateFileType[TemplateFileType["WikiPage"] = 1] = "WikiPage";
    TemplateFileType[TemplateFileType["FormPage"] = 2] = "FormPage";
    TemplateFileType[TemplateFileType["ClientSidePage"] = 3] = "ClientSidePage";
})(TemplateFileType || (TemplateFileType = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "qoFj":
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/site-designs/index.js ***!
  \****************************************************/
/*! exports provided: SiteDesigns */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "eDdy");
/* harmony import */ var _rest__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../rest */ "9l4K");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "w4d6");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteDesigns", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["SiteDesigns"]; });





Reflect.defineProperty(_rest__WEBPACK_IMPORTED_MODULE_1__["SPRest"].prototype, "siteDesigns", {
    configurable: true,
    enumerable: true,
    get: function () {
        return Object(_types__WEBPACK_IMPORTED_MODULE_2__["SiteDesigns"])(this._baseUrl);
    },
});
//# sourceMappingURL=index.js.map

/***/ }),

/***/ "t3D1":
/*!*****************************************************************!*\
  !*** ./node_modules/@pnp/odata/node_modules/tslib/tslib.es6.js ***!
  \*****************************************************************/
/*! exports provided: __extends, __assign, __rest, __decorate, __param, __metadata, __awaiter, __generator, __exportStar, __values, __read, __spread, __spreadArrays, __await, __asyncGenerator, __asyncDelegator, __asyncValues, __makeTemplateObject, __importStar, __importDefault */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__extends", function() { return __extends; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__assign", function() { return __assign; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__rest", function() { return __rest; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__decorate", function() { return __decorate; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__param", function() { return __param; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__metadata", function() { return __metadata; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__awaiter", function() { return __awaiter; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__generator", function() { return __generator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__exportStar", function() { return __exportStar; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__values", function() { return __values; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__read", function() { return __read; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__spread", function() { return __spread; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__spreadArrays", function() { return __spreadArrays; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__await", function() { return __await; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncGenerator", function() { return __asyncGenerator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncDelegator", function() { return __asyncDelegator; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__asyncValues", function() { return __asyncValues; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__makeTemplateObject", function() { return __makeTemplateObject; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__importStar", function() { return __importStar; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "__importDefault", function() { return __importDefault; });
/*! *****************************************************************************
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the Apache License, Version 2.0 (the "License"); you may not use
this file except in compliance with the License. You may obtain a copy of the
License at http://www.apache.org/licenses/LICENSE-2.0

THIS CODE IS PROVIDED ON AN *AS IS* BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
KIND, EITHER EXPRESS OR IMPLIED, INCLUDING WITHOUT LIMITATION ANY IMPLIED
WARRANTIES OR CONDITIONS OF TITLE, FITNESS FOR A PARTICULAR PURPOSE,
MERCHANTABLITY OR NON-INFRINGEMENT.

See the Apache Version 2.0 License for specific language governing permissions
and limitations under the License.
***************************************************************************** */
/* global Reflect, Promise */

var extendStatics = function(d, b) {
    extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return extendStatics(d, b);
};

function __extends(d, b) {
    extendStatics(d, b);
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
}

var __assign = function() {
    __assign = Object.assign || function __assign(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p)) t[p] = s[p];
        }
        return t;
    }
    return __assign.apply(this, arguments);
}

function __rest(s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
}

function __decorate(decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
}

function __param(paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
}

function __metadata(metadataKey, metadataValue) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(metadataKey, metadataValue);
}

function __awaiter(thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
}

function __generator(thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
}

function __exportStar(m, exports) {
    for (var p in m) if (!exports.hasOwnProperty(p)) exports[p] = m[p];
}

function __values(o) {
    var m = typeof Symbol === "function" && o[Symbol.iterator], i = 0;
    if (m) return m.call(o);
    return {
        next: function () {
            if (o && i >= o.length) o = void 0;
            return { value: o && o[i++], done: !o };
        }
    };
}

function __read(o, n) {
    var m = typeof Symbol === "function" && o[Symbol.iterator];
    if (!m) return o;
    var i = m.call(o), r, ar = [], e;
    try {
        while ((n === void 0 || n-- > 0) && !(r = i.next()).done) ar.push(r.value);
    }
    catch (error) { e = { error: error }; }
    finally {
        try {
            if (r && !r.done && (m = i["return"])) m.call(i);
        }
        finally { if (e) throw e.error; }
    }
    return ar;
}

function __spread() {
    for (var ar = [], i = 0; i < arguments.length; i++)
        ar = ar.concat(__read(arguments[i]));
    return ar;
}

function __spreadArrays() {
    for (var s = 0, i = 0, il = arguments.length; i < il; i++) s += arguments[i].length;
    for (var r = Array(s), k = 0, i = 0; i < il; i++)
        for (var a = arguments[i], j = 0, jl = a.length; j < jl; j++, k++)
            r[k] = a[j];
    return r;
};

function __await(v) {
    return this instanceof __await ? (this.v = v, this) : new __await(v);
}

function __asyncGenerator(thisArg, _arguments, generator) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var g = generator.apply(thisArg, _arguments || []), i, q = [];
    return i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i;
    function verb(n) { if (g[n]) i[n] = function (v) { return new Promise(function (a, b) { q.push([n, v, a, b]) > 1 || resume(n, v); }); }; }
    function resume(n, v) { try { step(g[n](v)); } catch (e) { settle(q[0][3], e); } }
    function step(r) { r.value instanceof __await ? Promise.resolve(r.value.v).then(fulfill, reject) : settle(q[0][2], r); }
    function fulfill(value) { resume("next", value); }
    function reject(value) { resume("throw", value); }
    function settle(f, v) { if (f(v), q.shift(), q.length) resume(q[0][0], q[0][1]); }
}

function __asyncDelegator(o) {
    var i, p;
    return i = {}, verb("next"), verb("throw", function (e) { throw e; }), verb("return"), i[Symbol.iterator] = function () { return this; }, i;
    function verb(n, f) { i[n] = o[n] ? function (v) { return (p = !p) ? { value: __await(o[n](v)), done: n === "return" } : f ? f(v) : v; } : f; }
}

function __asyncValues(o) {
    if (!Symbol.asyncIterator) throw new TypeError("Symbol.asyncIterator is not defined.");
    var m = o[Symbol.asyncIterator], i;
    return m ? m.call(o) : (o = typeof __values === "function" ? __values(o) : o[Symbol.iterator](), i = {}, verb("next"), verb("throw"), verb("return"), i[Symbol.asyncIterator] = function () { return this; }, i);
    function verb(n) { i[n] = o[n] && function (v) { return new Promise(function (resolve, reject) { v = o[n](v), settle(resolve, reject, v.done, v.value); }); }; }
    function settle(resolve, reject, d, v) { Promise.resolve(v).then(function(v) { resolve({ value: v, done: d }); }, reject); }
}

function __makeTemplateObject(cooked, raw) {
    if (Object.defineProperty) { Object.defineProperty(cooked, "raw", { value: raw }); } else { cooked.raw = raw; }
    return cooked;
};

function __importStar(mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result.default = mod;
    return result;
}

function __importDefault(mod) {
    return (mod && mod.__esModule) ? mod : { default: mod };
}


/***/ }),

/***/ "t8hP":
/*!******************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/toResourcePath.js ***!
  \******************************************************/
/*! exports provided: toResourcePath */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "toResourcePath", function() { return toResourcePath; });
function toResourcePath(url) {
    return {
        DecodedUrl: url,
        __metadata: { type: "SP.ResourcePath" },
    };
}
//# sourceMappingURL=toResourcePath.js.map

/***/ }),

/***/ "tCQJ":
/*!***************************************!*\
  !*** ./node_modules/@pnp/sp/types.js ***!
  \***************************************/
/*! exports provided: PrincipalType, PrincipalSource, PageType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PrincipalType", function() { return PrincipalType; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PrincipalSource", function() { return PrincipalSource; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "PageType", function() { return PageType; });
// reference: https://msdn.microsoft.com/en-us/library/office/dn600183.aspx
/**
 * Specifies the type of a principal.
 */
/* tslint:disable:no-bitwise */
var PrincipalType;
(function (PrincipalType) {
    /**
     * Enumeration whose value specifies no principal type.
     */
    PrincipalType[PrincipalType["None"] = 0] = "None";
    /**
     * Enumeration whose value specifies a user as the principal type.
     */
    PrincipalType[PrincipalType["User"] = 1] = "User";
    /**
     * Enumeration whose value specifies a distribution list as the principal type.
     */
    PrincipalType[PrincipalType["DistributionList"] = 2] = "DistributionList";
    /**
     * Enumeration whose value specifies a security group as the principal type.
     */
    PrincipalType[PrincipalType["SecurityGroup"] = 4] = "SecurityGroup";
    /**
     * Enumeration whose value specifies a group as the principal type.
     */
    PrincipalType[PrincipalType["SharePointGroup"] = 8] = "SharePointGroup";
    /**
     * Enumeration whose value specifies all principal types.
     */
    PrincipalType[PrincipalType["All"] = 15] = "All";
})(PrincipalType || (PrincipalType = {}));
/* tslint:enable:no-bitwise */
/**
 * Specifies the source of a principal.
 */
/* tslint:disable:no-bitwise */
var PrincipalSource;
(function (PrincipalSource) {
    /**
     * Enumeration whose value specifies no principal source.
     */
    PrincipalSource[PrincipalSource["None"] = 0] = "None";
    /**
     * Enumeration whose value specifies user information list as the principal source.
     */
    PrincipalSource[PrincipalSource["UserInfoList"] = 1] = "UserInfoList";
    /**
     * Enumeration whose value specifies Active Directory as the principal source.
     */
    PrincipalSource[PrincipalSource["Windows"] = 2] = "Windows";
    /**
     * Enumeration whose value specifies the current membership provider as the principal source.
     */
    PrincipalSource[PrincipalSource["MembershipProvider"] = 4] = "MembershipProvider";
    /**
     * Enumeration whose value specifies the current role provider as the principal source.
     */
    PrincipalSource[PrincipalSource["RoleProvider"] = 8] = "RoleProvider";
    /**
     * Enumeration whose value specifies all principal sources.
     */
    PrincipalSource[PrincipalSource["All"] = 15] = "All";
})(PrincipalSource || (PrincipalSource = {}));
var PageType;
(function (PageType) {
    PageType[PageType["Invalid"] = -1] = "Invalid";
    PageType[PageType["DefaultView"] = 0] = "DefaultView";
    PageType[PageType["NormalView"] = 1] = "NormalView";
    PageType[PageType["DialogView"] = 2] = "DialogView";
    PageType[PageType["View"] = 3] = "View";
    PageType[PageType["DisplayForm"] = 4] = "DisplayForm";
    PageType[PageType["DisplayFormDialog"] = 5] = "DisplayFormDialog";
    PageType[PageType["EditForm"] = 6] = "EditForm";
    PageType[PageType["EditFormDialog"] = 7] = "EditFormDialog";
    PageType[PageType["NewForm"] = 8] = "NewForm";
    PageType[PageType["NewFormDialog"] = 9] = "NewFormDialog";
    PageType[PageType["SolutionForm"] = 10] = "SolutionForm";
    PageType[PageType["PAGE_MAXITEMS"] = 11] = "PAGE_MAXITEMS";
})(PageType || (PageType = {}));
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "tSEd":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/extractweburl.js ***!
  \*****************************************************/
/*! exports provided: extractWebUrl */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "extractWebUrl", function() { return extractWebUrl; });
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/common */ "zqEK");

function extractWebUrl(candidateUrl) {
    if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["stringIsNullOrEmpty"])(candidateUrl)) {
        return "";
    }
    var index = candidateUrl.indexOf("_api/");
    if (index < 0) {
        index = candidateUrl.indexOf("_vti_bin/");
    }
    if (index > -1) {
        return candidateUrl.substr(0, index);
    }
    // if all else fails just give them what they gave us back
    return candidateUrl;
}
//# sourceMappingURL=extractweburl.js.map

/***/ }),

/***/ "uMGh":
/*!*********************************************************!*\
  !*** ./node_modules/@pnp/sp/user-custom-actions/web.js ***!
  \*********************************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "1mTr");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "userCustomActions", _types__WEBPACK_IMPORTED_MODULE_2__["UserCustomActions"]);
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "vX5r":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/site-users/index.js ***!
  \**************************************************/
/*! exports provided: SiteUser, SiteUsers */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "EjWy");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "y+KB");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteUser", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["SiteUser"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SiteUsers", function() { return _types__WEBPACK_IMPORTED_MODULE_1__["SiteUsers"]; });



//# sourceMappingURL=index.js.map

/***/ }),

/***/ "vXSC":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/security/funcs.js ***!
  \************************************************/
/*! exports provided: getUserEffectivePermissions, getCurrentUserEffectivePermissions, breakRoleInheritance, resetRoleInheritance, userHasPermissions, currentUserHasPermissions, hasPermissions */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getUserEffectivePermissions", function() { return getUserEffectivePermissions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getCurrentUserEffectivePermissions", function() { return getCurrentUserEffectivePermissions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "breakRoleInheritance", function() { return breakRoleInheritance; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "resetRoleInheritance", function() { return resetRoleInheritance; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "userHasPermissions", function() { return userHasPermissions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "currentUserHasPermissions", function() { return currentUserHasPermissions; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "hasPermissions", function() { return hasPermissions; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "1dmT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../operations */ "UK2s");





/**
* Gets the effective permissions for the user supplied
*
* @param loginName The claims username for the user (ex: i:0#.f|membership|user@domain.com)
*/
function getUserEffectivePermissions(loginName) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var q, r;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0:
                    q = this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryableInstance"], "getUserEffectivePermissions(@user)");
                    q.query.set("@user", "'" + encodeURIComponent(loginName) + "'");
                    return [4 /*yield*/, q.get()];
                case 1:
                    r = _a.sent();
                    // handle verbose mode
                    return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["hOP"])(r, "GetUserEffectivePermissions") ? r.GetUserEffectivePermissions : r];
            }
        });
    });
}
/**
 * Gets the effective permissions for the current user
 */
function getCurrentUserEffectivePermissions() {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var q;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            q = this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryable"], "EffectiveBasePermissions");
            return [2 /*return*/, q.get().then(function (r) {
                    // handle verbose mode
                    return Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["hOP"])(r, "EffectiveBasePermissions") ? r.EffectiveBasePermissions : r;
                })];
        });
    });
}
/**
 * Breaks the security inheritance at this level optinally copying permissions and clearing subscopes
 *
 * @param copyRoleAssignments If true the permissions are copied from the current parent scope
 * @param clearSubscopes Optional. true to make all child securable objects inherit role assignments from the current object
 */
function breakRoleInheritance(copyRoleAssignments, clearSubscopes) {
    if (copyRoleAssignments === void 0) { copyRoleAssignments = false; }
    if (clearSubscopes === void 0) { clearSubscopes = false; }
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryable"], "breakroleinheritance(copyroleassignments=" + copyRoleAssignments + ", clearsubscopes=" + clearSubscopes + ")"))];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
/**
 * Removes the local role assignments so that it re-inherit role assignments from the parent object.
 *
 */
function resetRoleInheritance() {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(this.clone(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["SharePointQueryable"], "resetroleinheritance"))];
                case 1:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    });
}
/**
 * Determines if a given user has the appropriate permissions
 *
 * @param loginName The user to check
 * @param permission The permission being checked
 */
function userHasPermissions(loginName, permission) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var perms;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getUserEffectivePermissions.call(this, loginName)];
                case 1:
                    perms = _a.sent();
                    return [2 /*return*/, this.hasPermissions(perms, permission)];
            }
        });
    });
}
/**
 * Determines if the current user has the requested permissions
 *
 * @param permission The permission we wish to check
 */
function currentUserHasPermissions(permission) {
    return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
        var perms;
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, getCurrentUserEffectivePermissions.call(this)];
                case 1:
                    perms = _a.sent();
                    return [2 /*return*/, this.hasPermissions(perms, permission)];
            }
        });
    });
}
/**
 * Taken from sp.js, checks the supplied permissions against the mask
 *
 * @param value The security principal's permissions on the given object
 * @param perm The permission checked against the value
 */
/* tslint:disable:no-bitwise */
function hasPermissions(value, perm) {
    if (!perm) {
        return true;
    }
    if (perm === _types__WEBPACK_IMPORTED_MODULE_1__["PermissionKind"].FullMask) {
        return (value.High & 32767) === 32767 && value.Low === 65535;
    }
    perm = perm - 1;
    var num = 1;
    if (perm >= 0 && perm < 32) {
        num = num << perm;
        return 0 !== (value.Low & num);
    }
    else if (perm >= 32 && perm < 64) {
        num = num << perm - 32;
        return 0 !== (value.High & num);
    }
    return false;
}
/* tslint:enable */
//# sourceMappingURL=funcs.js.map

/***/ }),

/***/ "vttn":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/list.js ***!
  \**********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _lists_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../lists/types */ "hy0S");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "4tXH");



Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_lists_types__WEBPACK_IMPORTED_MODULE_1__["_List"], "rootFolder", _types__WEBPACK_IMPORTED_MODULE_2__["Folder"], "rootFolder");
//# sourceMappingURL=list.js.map

/***/ }),

/***/ "w4d6":
/*!****************************************************!*\
  !*** ./node_modules/@pnp/sp/site-designs/types.js ***!
  \****************************************************/
/*! exports provided: _SiteDesigns, SiteDesigns */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SiteDesigns", function() { return _SiteDesigns; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SiteDesigns", function() { return SiteDesigns; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _utils_extractweburl__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../utils/extractweburl */ "tSEd");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../telemetry */ "h3gW");







var _SiteDesigns = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SiteDesigns, _super);
    function _SiteDesigns(baseUrl, methodName) {
        if (methodName === void 0) { methodName = ""; }
        var _this = this;
        var url = typeof baseUrl === "string" ? baseUrl : baseUrl.toUrl();
        _this = _super.call(this, Object(_utils_extractweburl__WEBPACK_IMPORTED_MODULE_2__["extractWebUrl"])(url), "_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + methodName) || this;
        return _this;
    }
    _SiteDesigns.prototype.execute = function (props) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_4__["spPost"])(this, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_3__["body"])(props, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_3__["headers"])({ "Content-Type": "application/json;charset=utf-8" })));
    };
    /**
     * Creates a new site design available to users when they create a new site from the SharePoint home page.
     *
     * @param creationInfo A sitedesign creation information object
     */
    _SiteDesigns.prototype.createSiteDesign = function (creationInfo) {
        return this.clone(SiteDesignsCloneFactory, "CreateSiteDesign").execute({ info: creationInfo });
    };
    /**
     * Applies a site design to an existing site collection.
     *
     * @param siteDesignId The ID of the site design to apply.
     * @param webUrl The URL of the site collection where you want to apply the site design.
     */
    _SiteDesigns.prototype.applySiteDesign = function (siteDesignId, webUrl) {
        return this.clone(SiteDesignsCloneFactory, "ApplySiteDesign").execute({ siteDesignId: siteDesignId, "webUrl": webUrl });
    };
    /**
     * Gets the list of available site designs
     */
    _SiteDesigns.prototype.getSiteDesigns = function () {
        return this.clone(SiteDesignsCloneFactory, "GetSiteDesigns").execute({});
    };
    /**
     * Gets information about a specific site design.
     * @param id The ID of the site design to get information about.
     */
    _SiteDesigns.prototype.getSiteDesignMetadata = function (id) {
        return this.clone(SiteDesignsCloneFactory, "GetSiteDesignMetadata").execute({ id: id });
    };
    /**
     * Updates a site design with new values. In the REST call, all parameters are optional except the site script Id.
     * If you had previously set the IsDefault parameter to TRUE and wish it to remain true, you must pass in this parameter again (otherwise it will be reset to FALSE).
     * @param updateInfo A sitedesign update information object
     */
    _SiteDesigns.prototype.updateSiteDesign = function (updateInfo) {
        return this.clone(SiteDesignsCloneFactory, "UpdateSiteDesign").execute({ updateInfo: updateInfo });
    };
    /**
     * Deletes a site design.
     * @param id The ID of the site design to delete.
     */
    _SiteDesigns.prototype.deleteSiteDesign = function (id) {
        return this.clone(SiteDesignsCloneFactory, "DeleteSiteDesign").execute({ id: id });
    };
    /**
     * Gets a list of principals that have access to a site design.
     * @param id The ID of the site design to get rights information from.
     */
    _SiteDesigns.prototype.getSiteDesignRights = function (id) {
        return this.clone(SiteDesignsCloneFactory, "GetSiteDesignRights").execute({ id: id });
    };
    /**
     * Grants access to a site design for one or more principals.
     * @param id The ID of the site design to grant rights on.
     * @param principalNames An array of one or more principals to grant view rights.
     *                       Principals can be users or mail-enabled security groups in the form of "alias" or "alias@<domain name>.com"
     * @param grantedRights Always set to 1. This represents the View right.
     */
    _SiteDesigns.prototype.grantSiteDesignRights = function (id, principalNames, grantedRights) {
        if (grantedRights === void 0) { grantedRights = 1; }
        return this.clone(SiteDesignsCloneFactory, "GrantSiteDesignRights")
            .execute({
            "grantedRights": grantedRights.toString(),
            "id": id,
            "principalNames": principalNames,
        });
    };
    /**
     * Revokes access from a site design for one or more principals.
     * @param id The ID of the site design to revoke rights from.
     * @param principalNames An array of one or more principals to revoke view rights from.
     *                       If all principals have rights revoked on the site design, the site design becomes viewable to everyone.
     */
    _SiteDesigns.prototype.revokeSiteDesignRights = function (id, principalNames) {
        return this.clone(SiteDesignsCloneFactory, "RevokeSiteDesignRights")
            .execute({
            "id": id,
            "principalNames": principalNames,
        });
    };
    /**
     * Adds a site design task on the specified web url to be invoked asynchronously.
     * @param webUrl The absolute url of the web on where to create the task
     * @param siteDesignId The ID of the site design to create a task for
     */
    _SiteDesigns.prototype.addSiteDesignTask = function (webUrl, siteDesignId) {
        return this.clone(SiteDesignsCloneFactory, "AddSiteDesignTask")
            .execute({ "webUrl": webUrl, "siteDesignId": siteDesignId });
    };
    /**
     * Adds a site design task on the current web to be invoked asynchronously.
     * @param siteDesignId The ID of the site design to create a task for
     */
    _SiteDesigns.prototype.addSiteDesignTaskToCurrentWeb = function (siteDesignId) {
        return this.clone(SiteDesignsCloneFactory, "AddSiteDesignTaskToCurrentWeb")
            .execute({ "siteDesignId": siteDesignId });
    };
    /**
     * Retrieves the site design task, if the task has finished running null will be returned
     * @param id The ID of the site design task
     */
    _SiteDesigns.prototype.getSiteDesignTask = function (id) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var task;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.clone(SiteDesignsCloneFactory, "GetSiteDesignTask")
                            .execute({ "taskId": id })];
                    case 1:
                        task = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_5__["hOP"])(task, "ID") ? task : null];
                }
            });
        });
    };
    /**
     * Retrieves a list of site design that have run on a specific web
     * @param webUrl The url of the web where the site design was applied
     * @param siteDesignId (Optional) the site design ID, if not provided will return all site design runs
     */
    _SiteDesigns.prototype.getSiteDesignRun = function (webUrl, siteDesignId) {
        return this.clone(SiteDesignsCloneFactory, "GetSiteDesignRun")
            .execute({ "webUrl": webUrl, siteDesignId: siteDesignId });
    };
    /**
     * Retrieves the status of a site design that has been run or is still running
     * @param webUrl The url of the web where the site design was applied
     * @param runId the run ID
     */
    _SiteDesigns.prototype.getSiteDesignRunStatus = function (webUrl, runId) {
        return this.clone(SiteDesignsCloneFactory, "GetSiteDesignRunStatus")
            .execute({ "webUrl": webUrl, runId: runId });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.createSiteDesign")
    ], _SiteDesigns.prototype, "createSiteDesign", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.applySiteDesign")
    ], _SiteDesigns.prototype, "applySiteDesign", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.getSiteDesigns")
    ], _SiteDesigns.prototype, "getSiteDesigns", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.getSiteDesignMetadata")
    ], _SiteDesigns.prototype, "getSiteDesignMetadata", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.updateSiteDesign")
    ], _SiteDesigns.prototype, "updateSiteDesign", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.deleteSiteDesign")
    ], _SiteDesigns.prototype, "deleteSiteDesign", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.getSiteDesignRights")
    ], _SiteDesigns.prototype, "getSiteDesignRights", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.grantSiteDesignRights")
    ], _SiteDesigns.prototype, "grantSiteDesignRights", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.revokeSiteDesignRights")
    ], _SiteDesigns.prototype, "revokeSiteDesignRights", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.addSiteDesignTask")
    ], _SiteDesigns.prototype, "addSiteDesignTask", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.addSiteDesignTaskToCurrentWeb")
    ], _SiteDesigns.prototype, "addSiteDesignTaskToCurrentWeb", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.getSiteDesignTask")
    ], _SiteDesigns.prototype, "getSiteDesignTask", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.getSiteDesignRun")
    ], _SiteDesigns.prototype, "getSiteDesignRun", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_6__["tag"])("sd.getSiteDesignRunStatus")
    ], _SiteDesigns.prototype, "getSiteDesignRunStatus", null);
    return _SiteDesigns;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryable"]));

var SiteDesigns = function (baseUrl, methodName) { return new _SiteDesigns(baseUrl, methodName); };
var SiteDesignsCloneFactory = function (baseUrl, methodName) {
    if (methodName === void 0) { methodName = ""; }
    return SiteDesigns(baseUrl, methodName);
};
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "wMck":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/comments/types.js ***!
  \************************************************/
/*! exports provided: _Comments, Comments, _Comment, Comment, _Replies, Replies */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Comments", function() { return _Comments; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Comments", function() { return Comments; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Comment", function() { return _Comment; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Comment", function() { return Comment; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Replies", function() { return _Replies; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Replies", function() { return Replies; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _odata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../odata */ "5dQz");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../telemetry */ "h3gW");









var _Comments = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Comments, _super);
    function _Comments() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Adds a new comment to this collection
     *
     * @param info Comment information to add
     */
    _Comments.prototype.add = function (info) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, d;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (typeof info === "string") {
                            info = { text: info };
                        }
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_6__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("Microsoft.SharePoint.Comments.comment"), info));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Comments, null), postBody)];
                    case 1:
                        d = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["assign"])(this.getById(d.id), d)];
                }
            });
        });
    };
    /**
     * Gets a comment by id
     *
     * @param id Id of the comment to load
     */
    _Comments.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(Comment(this).concat("(" + id + ")"), "coms.getById");
    };
    /**
     * Deletes all the comments in this collection
     */
    _Comments.prototype.clear = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(this.clone(Comments, "DeleteAll"), "coms.clear"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("coms.add")
    ], _Comments.prototype, "add", null);
    _Comments = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_1__["defaultPath"])("comments")
    ], _Comments);
    return _Comments;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var Comments = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_Comments);
var _Comment = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Comment, _super);
    function _Comment() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Object.defineProperty(_Comment.prototype, "replies", {
        /**
         * A comment's replies
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(Replies(this), "com.replies");
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Likes the comment as the current user
     */
    _Comment.prototype.like = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Comment, "Like"));
    };
    /**
     * Unlikes the comment as the current user
     */
    _Comment.prototype.unlike = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Comment, "Unlike"));
    };
    /**
     * Deletes this comment
     */
    _Comment.prototype.delete = function () {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Comment, "DeleteComment"));
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("com.like")
    ], _Comment.prototype, "like", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("com.unlike")
    ], _Comment.prototype, "unlike", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("com.delete")
    ], _Comment.prototype, "delete", null);
    return _Comment;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableInstance"]));

var Comment = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_Comment);
var _Replies = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Replies, _super);
    function _Replies() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Adds a new reply to this collection
     *
     * @param info Comment information to add
     */
    _Replies.prototype.add = function (info) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var postBody, d;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (typeof info === "string") {
                            info = { text: info };
                        }
                        postBody = Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_6__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_5__["metadata"])("Microsoft.SharePoint.Comments.comment"), info));
                        return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(Replies, null), postBody)];
                    case 1:
                        d = _a.sent();
                        return [2 /*return*/, Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["assign"])(Comment(Object(_odata__WEBPACK_IMPORTED_MODULE_4__["odataUrlFrom"])(d)), d)];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("reps.add")
    ], _Replies.prototype, "add", null);
    _Replies = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_1__["defaultPath"])("replies")
    ], _Replies);
    return _Replies;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["_SharePointQueryableCollection"]));

var Replies = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_2__["spInvokableFactory"])(_Replies);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "wdEC":
/*!****************************************************************************!*\
  !*** ./lib/extensions/announcements/AnnouncementsApplicationCustomizer.js ***!
  \****************************************************************************/
/*! exports provided: QUALIFIED_NAME, default */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "QUALIFIED_NAME", function() { return QUALIFIED_NAME; });
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ "cDcd");
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var react_dom__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! react-dom */ "faye");
/* harmony import */ var react_dom__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(react_dom__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_decorators__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/decorators */ "wxtz");
/* harmony import */ var _microsoft_decorators__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_decorators__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @microsoft/sp-core-library */ "UWqr");
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var _microsoft_sp_application_base__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @microsoft/sp-application-base */ "GPet");
/* harmony import */ var _microsoft_sp_application_base__WEBPACK_IMPORTED_MODULE_4___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_application_base__WEBPACK_IMPORTED_MODULE_4__);
/* harmony import */ var _pnp_sp_presets_all__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/sp/presets/all */ "F/m7");
/* harmony import */ var announcementsStrings__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! announcementsStrings */ "+n6P");
/* harmony import */ var announcementsStrings__WEBPACK_IMPORTED_MODULE_6___default = /*#__PURE__*/__webpack_require__.n(announcementsStrings__WEBPACK_IMPORTED_MODULE_6__);
/* harmony import */ var _components_Announcements__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./components/Announcements */ "pM8y");
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (undefined && undefined.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};








var QUALIFIED_NAME = 'Extension.ApplicationCustomizer.Announcements';
var AnnouncementsApplicationCustomizer = /** @class */ (function (_super) {
    __extends(AnnouncementsApplicationCustomizer, _super);
    function AnnouncementsApplicationCustomizer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    AnnouncementsApplicationCustomizer.prototype.onInit = function () {
        return __awaiter(this, void 0, void 0, function () {
            var e, header, error, site, tenantUrl, elem;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.onInit.call(this)];
                    case 1:
                        _a.sent();
                        _pnp_sp_presets_all__WEBPACK_IMPORTED_MODULE_5__["sp"].setup(this.context);
                        _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_3__["Log"].info(QUALIFIED_NAME, "Initializing " + announcementsStrings__WEBPACK_IMPORTED_MODULE_6__["Title"]);
                        if (!this.properties.siteUrl || !this.properties.listName) {
                            e = new Error('Missing required configuration parameters');
                            _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_3__["Log"].error(QUALIFIED_NAME, e);
                            return [2 /*return*/, Promise.reject(e)];
                        }
                        header = this.context.placeholderProvider.tryCreateContent(_microsoft_sp_application_base__WEBPACK_IMPORTED_MODULE_4__["PlaceholderName"].Top);
                        if (!header) {
                            error = new Error('Could not find placeholder Top');
                            _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_3__["Log"].error(QUALIFIED_NAME, error);
                            return [2 /*return*/, Promise.reject(error)];
                        }
                        site = this.context.pageContext.site;
                        tenantUrl = site.absoluteUrl.replace(site.serverRelativeUrl, "");
                        elem = react__WEBPACK_IMPORTED_MODULE_0__["createElement"](_components_Announcements__WEBPACK_IMPORTED_MODULE_7__["default"], {
                            siteUrl: "" + tenantUrl + this.properties.siteUrl,
                            listName: this.properties.listName,
                            culture: this.context.pageContext.cultureInfo.currentUICultureName
                        });
                        react_dom__WEBPACK_IMPORTED_MODULE_1__["render"](elem, header.domElement);
                        return [2 /*return*/, Promise.resolve()];
                }
            });
        });
    };
    __decorate([
        _microsoft_decorators__WEBPACK_IMPORTED_MODULE_2__["override"]
    ], AnnouncementsApplicationCustomizer.prototype, "onInit", null);
    return AnnouncementsApplicationCustomizer;
}(_microsoft_sp_application_base__WEBPACK_IMPORTED_MODULE_4__["BaseApplicationCustomizer"]));
/* harmony default export */ __webpack_exports__["default"] = (AnnouncementsApplicationCustomizer);


/***/ }),

/***/ "wmuB":
/*!**********************************************!*\
  !*** ./node_modules/@pnp/sp/fields/index.js ***!
  \**********************************************/
/*! exports provided: Fields, Field, AddFieldOptions, CalendarType, ChoiceFieldFormatType, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType, FieldTypes, FieldUserSelectionMode, UrlFieldFormatType */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _web__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./web */ "psVa");
/* harmony import */ var _list__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./list */ "6k8n");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "KYlJ");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Fields", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["Fields"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Field", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["Field"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "AddFieldOptions", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["AddFieldOptions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CalendarType", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["CalendarType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ChoiceFieldFormatType", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["ChoiceFieldFormatType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "DateTimeFieldFormatType", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["DateTimeFieldFormatType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "DateTimeFieldFriendlyFormatType", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["DateTimeFieldFriendlyFormatType"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FieldTypes", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["FieldTypes"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FieldUserSelectionMode", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["FieldUserSelectionMode"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "UrlFieldFormatType", function() { return _types__WEBPACK_IMPORTED_MODULE_2__["UrlFieldFormatType"]; });





//# sourceMappingURL=index.js.map

/***/ }),

/***/ "wxtz":
/*!****************************************!*\
  !*** external "@microsoft/decorators" ***!
  \****************************************/
/*! no static exports found */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_wxtz__;

/***/ }),

/***/ "xT8B":
/*!******************************************!*\
  !*** ./node_modules/@pnp/odata/index.js ***!
  \******************************************/
/*! exports provided: extendGlobal, extendObj, extendFactory, clearGlobalExtensions, enableExtensions, disableExtensions, Batch, CachingOptions, CachingParserWrapper, addProp, invokableFactory, pipelineBinder, defaultPipelineBinder, ODataParser, TextParser, BlobParser, JSONParser, BufferParser, LambdaParser, HttpRequestError, setResult, pipe, requestPipelineMethod, PipelineMethods, getDefaultPipeline, cloneQueryableData, Queryable, body, headers */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _batch__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./batch */ "SQg3");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Batch", function() { return _batch__WEBPACK_IMPORTED_MODULE_0__["Batch"]; });

/* harmony import */ var _caching__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./caching */ "Zl34");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CachingOptions", function() { return _caching__WEBPACK_IMPORTED_MODULE_1__["CachingOptions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "CachingParserWrapper", function() { return _caching__WEBPACK_IMPORTED_MODULE_1__["CachingParserWrapper"]; });

/* harmony import */ var _add_prop__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./add-prop */ "W7+7");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "addProp", function() { return _add_prop__WEBPACK_IMPORTED_MODULE_2__["addProp"]; });

/* harmony import */ var _invokable_binder__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./invokable-binder */ "pG6f");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "invokableFactory", function() { return _invokable_binder__WEBPACK_IMPORTED_MODULE_3__["invokableFactory"]; });

/* harmony import */ var _pipeline_binder__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./pipeline-binder */ "5e8D");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "pipelineBinder", function() { return _pipeline_binder__WEBPACK_IMPORTED_MODULE_4__["pipelineBinder"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "defaultPipelineBinder", function() { return _pipeline_binder__WEBPACK_IMPORTED_MODULE_4__["defaultPipelineBinder"]; });

/* harmony import */ var _parsers__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./parsers */ "xbXq");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "ODataParser", function() { return _parsers__WEBPACK_IMPORTED_MODULE_5__["ODataParser"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "TextParser", function() { return _parsers__WEBPACK_IMPORTED_MODULE_5__["TextParser"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "BlobParser", function() { return _parsers__WEBPACK_IMPORTED_MODULE_5__["BlobParser"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "JSONParser", function() { return _parsers__WEBPACK_IMPORTED_MODULE_5__["JSONParser"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "BufferParser", function() { return _parsers__WEBPACK_IMPORTED_MODULE_5__["BufferParser"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "LambdaParser", function() { return _parsers__WEBPACK_IMPORTED_MODULE_5__["LambdaParser"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "HttpRequestError", function() { return _parsers__WEBPACK_IMPORTED_MODULE_5__["HttpRequestError"]; });

/* harmony import */ var _pipeline__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./pipeline */ "TZA7");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "setResult", function() { return _pipeline__WEBPACK_IMPORTED_MODULE_6__["setResult"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "pipe", function() { return _pipeline__WEBPACK_IMPORTED_MODULE_6__["pipe"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "requestPipelineMethod", function() { return _pipeline__WEBPACK_IMPORTED_MODULE_6__["requestPipelineMethod"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PipelineMethods", function() { return _pipeline__WEBPACK_IMPORTED_MODULE_6__["PipelineMethods"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "getDefaultPipeline", function() { return _pipeline__WEBPACK_IMPORTED_MODULE_6__["getDefaultPipeline"]; });

/* harmony import */ var _queryable__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./queryable */ "dKLh");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "cloneQueryableData", function() { return _queryable__WEBPACK_IMPORTED_MODULE_7__["cloneQueryableData"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "Queryable", function() { return _queryable__WEBPACK_IMPORTED_MODULE_7__["Queryable"]; });

/* harmony import */ var _request_builders__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./request-builders */ "xsXS");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "body", function() { return _request_builders__WEBPACK_IMPORTED_MODULE_8__["body"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "headers", function() { return _request_builders__WEBPACK_IMPORTED_MODULE_8__["headers"]; });

/* harmony import */ var _invokable_extensions__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! ./invokable-extensions */ "01TU");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "extendGlobal", function() { return _invokable_extensions__WEBPACK_IMPORTED_MODULE_9__["extendGlobal"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "extendObj", function() { return _invokable_extensions__WEBPACK_IMPORTED_MODULE_9__["extendObj"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "extendFactory", function() { return _invokable_extensions__WEBPACK_IMPORTED_MODULE_9__["extendFactory"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "clearGlobalExtensions", function() { return _invokable_extensions__WEBPACK_IMPORTED_MODULE_9__["clearGlobalExtensions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "enableExtensions", function() { return _invokable_extensions__WEBPACK_IMPORTED_MODULE_9__["enableExtensions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "disableExtensions", function() { return _invokable_extensions__WEBPACK_IMPORTED_MODULE_9__["disableExtensions"]; });











//# sourceMappingURL=index.js.map

/***/ }),

/***/ "xbXq":
/*!********************************************!*\
  !*** ./node_modules/@pnp/odata/parsers.js ***!
  \********************************************/
/*! exports provided: ODataParser, TextParser, BlobParser, JSONParser, BufferParser, LambdaParser, HttpRequestError */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ODataParser", function() { return ODataParser; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "TextParser", function() { return TextParser; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "BlobParser", function() { return BlobParser; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "JSONParser", function() { return JSONParser; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "BufferParser", function() { return BufferParser; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "LambdaParser", function() { return LambdaParser; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "HttpRequestError", function() { return HttpRequestError; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "t3D1");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @pnp/common */ "zqEK");


var ODataParser = /** @class */ (function () {
    function ODataParser() {
    }
    ODataParser.prototype.parse = function (r) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            if (_this.handleError(r, reject)) {
                _this.parseImpl(r, resolve, reject);
            }
        });
    };
    ODataParser.prototype.parseImpl = function (r, resolve, reject) {
        var _this = this;
        if ((r.headers.has("Content-Length") && parseFloat(r.headers.get("Content-Length")) === 0) || r.status === 204) {
            resolve({});
        }
        else {
            // patch to handle cases of 200 response with no or whitespace only bodies (#487 & #545)
            r.text()
                .then(function (txt) { return txt.replace(/\s/ig, "").length > 0 ? JSON.parse(txt) : {}; })
                .then(function (json) { return resolve(_this.parseODataJSON(json)); })
                .catch(function (e) { return reject(e); });
        }
    };
    /**
     * Handles a response with ok === false by parsing the body and creating a ProcessHttpClientResponseException
     * which is passed to the reject delegate. This method returns true if there is no error, otherwise false
     *
     * @param r Current response object
     * @param reject reject delegate for the surrounding promise
     */
    ODataParser.prototype.handleError = function (r, reject) {
        if (!r.ok) {
            HttpRequestError.init(r).then(reject);
        }
        return r.ok;
    };
    /**
     * Normalizes the json response by removing the various nested levels
     *
     * @param json json object to parse
     */
    ODataParser.prototype.parseODataJSON = function (json) {
        var result = json;
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(json, "d")) {
            if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(json.d, "results")) {
                result = json.d.results;
            }
            else {
                result = json.d;
            }
        }
        else if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["hOP"])(json, "value")) {
            result = json.value;
        }
        return result;
    };
    return ODataParser;
}());

var TextParser = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(TextParser, _super);
    function TextParser() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    TextParser.prototype.parseImpl = function (r, resolve) {
        r.text().then(resolve);
    };
    return TextParser;
}(ODataParser));

var BlobParser = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(BlobParser, _super);
    function BlobParser() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    BlobParser.prototype.parseImpl = function (r, resolve) {
        r.blob().then(resolve);
    };
    return BlobParser;
}(ODataParser));

var JSONParser = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(JSONParser, _super);
    function JSONParser() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    JSONParser.prototype.parseImpl = function (r, resolve) {
        r.json().then(resolve);
    };
    return JSONParser;
}(ODataParser));

var BufferParser = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(BufferParser, _super);
    function BufferParser() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    BufferParser.prototype.parseImpl = function (r, resolve) {
        if (Object(_pnp_common__WEBPACK_IMPORTED_MODULE_1__["isFunc"])(r.arrayBuffer)) {
            r.arrayBuffer().then(resolve);
        }
        else {
            r.buffer().then(resolve);
        }
    };
    return BufferParser;
}(ODataParser));

var LambdaParser = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(LambdaParser, _super);
    function LambdaParser(parser) {
        var _this = _super.call(this) || this;
        _this.parser = parser;
        return _this;
    }
    LambdaParser.prototype.parseImpl = function (r, resolve) {
        this.parser(r).then(resolve);
    };
    return LambdaParser;
}(ODataParser));

var HttpRequestError = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(HttpRequestError, _super);
    function HttpRequestError(message, response, status, statusText) {
        if (status === void 0) { status = response.status; }
        if (statusText === void 0) { statusText = response.statusText; }
        var _this = _super.call(this, message) || this;
        _this.response = response;
        _this.status = status;
        _this.statusText = statusText;
        _this.isHttpRequestError = true;
        return _this;
    }
    HttpRequestError.init = function (r) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            var t;
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, r.clone().text()];
                    case 1:
                        t = _a.sent();
                        return [2 /*return*/, new HttpRequestError("Error making HttpClient request in queryable [" + r.status + "] " + r.statusText + " ::> " + t, r.clone())];
                }
            });
        });
    };
    return HttpRequestError;
}(Error));

//# sourceMappingURL=parsers.js.map

/***/ }),

/***/ "xsXS":
/*!*****************************************************!*\
  !*** ./node_modules/@pnp/odata/request-builders.js ***!
  \*****************************************************/
/*! exports provided: body, headers */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "body", function() { return body; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "headers", function() { return headers; });
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/common */ "zqEK");

function body(o, previous) {
    return Object.assign({ body: Object(_pnp_common__WEBPACK_IMPORTED_MODULE_0__["jsS"])(o) }, previous);
}
function headers(o, previous) {
    return Object.assign({ headers: o }, previous);
}
//# sourceMappingURL=request-builders.js.map

/***/ }),

/***/ "y+KB":
/*!**************************************************!*\
  !*** ./node_modules/@pnp/sp/site-users/types.js ***!
  \**************************************************/
/*! exports provided: _SiteUsers, SiteUsers, _SiteUser, SiteUser */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SiteUsers", function() { return _SiteUsers; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SiteUsers", function() { return SiteUsers; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_SiteUser", function() { return _SiteUser; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SiteUser", function() { return SiteUser; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _site_groups_types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../site-groups/types */ "UuUm");
/* harmony import */ var _pnp_common__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @pnp/common */ "zqEK");
/* harmony import */ var _utils_metadata__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ../utils/metadata */ "z067");
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _operations__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ../operations */ "UK2s");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ../telemetry */ "h3gW");









var _SiteUsers = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SiteUsers, _super);
    function _SiteUsers() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a user from the collection by id
     *
     * @param id The id of the user to retrieve
     */
    _SiteUsers.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(SiteUser(this, "getById(" + id + ")"), "sus.getById");
    };
    /**
     * Gets a user from the collection by email
     *
     * @param email The email address of the user to retrieve
     */
    _SiteUsers.prototype.getByEmail = function (email) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(SiteUser(this, "getByEmail('" + email + "')"), "sus.getByEmail");
    };
    /**
     * Gets a user from the collection by login name
     *
     * @param loginName The login name of the user to retrieve
     */
    _SiteUsers.prototype.getByLoginName = function (loginName) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(SiteUser(this).concat("('!@v::" + encodeURIComponent(loginName) + "')"), "sus.getByLoginName");
    };
    /**
     * Removes a user from the collection by id
     *
     * @param id The id of the user to remove
     */
    _SiteUsers.prototype.removeById = function (id) {
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this.clone(SiteUsers, "removeById(" + id + ")"));
    };
    /**
     * Removes a user from the collection by login name
     *
     * @param loginName The login name of the user to remove
     */
    _SiteUsers.prototype.removeByLoginName = function (loginName) {
        var o = this.clone(SiteUsers, "removeByLoginName(@v)");
        o.query.set("@v", "'" + encodeURIComponent(loginName) + "'");
        return Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(o);
    };
    /**
     * Adds a user to a site collection
     *
     * @param loginName The login name of the user to add  to a site collection
     *
     */
    _SiteUsers.prototype.add = function (loginName) {
        return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__awaiter"])(this, void 0, void 0, function () {
            return Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__generator"])(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, Object(_operations__WEBPACK_IMPORTED_MODULE_7__["spPost"])(this, Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_5__["body"])(Object(_pnp_common__WEBPACK_IMPORTED_MODULE_3__["assign"])(Object(_utils_metadata__WEBPACK_IMPORTED_MODULE_4__["metadata"])("SP.User"), { LoginName: loginName })))];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, this.getByLoginName(loginName)];
                }
            });
        });
    };
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("sus.remId")
    ], _SiteUsers.prototype, "removeById", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("sus.remLoginName")
    ], _SiteUsers.prototype, "removeByLoginName", null);
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"])("sus.add")
    ], _SiteUsers.prototype, "add", null);
    _SiteUsers = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_6__["defaultPath"])("siteusers")
    ], _SiteUsers);
    return _SiteUsers;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var SiteUsers = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_SiteUsers);
/**
 * Describes a single user
 *
 */
var _SiteUser = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_SiteUser, _super);
    function _SiteUser() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.delete = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["deleteable"])("su");
        /**
        * Updates this user instance with the supplied properties
        *
        * @param properties A plain object of property names and values to update for the user
        */
        _this.update = _this._update("SP.User", function (data) { return ({ data: data, user: _this }); });
        return _this;
    }
    Object.defineProperty(_SiteUser.prototype, "groups", {
        /**
         * Gets the groups for this user
         *
         */
        get: function () {
            return _telemetry__WEBPACK_IMPORTED_MODULE_8__["tag"].configure(Object(_site_groups_types__WEBPACK_IMPORTED_MODULE_2__["SiteGroups"])(this, "groups"), "su.groups");
        },
        enumerable: true,
        configurable: true
    });
    return _SiteUser;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var SiteUser = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_SiteUser);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "yGnx":
/*!***********************************************!*\
  !*** ./node_modules/@pnp/sp/webparts/file.js ***!
  \***********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _files_types__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../files/types */ "qfcF");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./types */ "pXI8");


_files_types__WEBPACK_IMPORTED_MODULE_0__["_File"].prototype.getLimitedWebPartManager = function (scope) {
    if (scope === void 0) { scope = _types__WEBPACK_IMPORTED_MODULE_1__["WebPartsPersonalizationScope"].Shared; }
    return Object(_types__WEBPACK_IMPORTED_MODULE_1__["LimitedWebPartManager"])(this, "getLimitedWebPartManager(scope=" + scope + ")");
};
//# sourceMappingURL=file.js.map

/***/ }),

/***/ "yLpj":
/*!***********************************!*\
  !*** (webpack)/buildin/global.js ***!
  \***********************************/
/*! no static exports found */
/***/ (function(module, exports) {

var g;

// This works in non-strict mode
g = (function() {
	return this;
})();

try {
	// This works if eval is allowed (see CSP)
	g = g || new Function("return this")();
} catch (e) {
	// This works if the window reference is available
	if (typeof window === "object") g = window;
}

// g can still be undefined, but nothing to do about it...
// We return undefined, instead of nothing here, so it's
// easier to handle this case. if(!global) { ...}

module.exports = g;


/***/ }),

/***/ "ymW/":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/folders/web.js ***!
  \*********************************************/
/*! no exports provided */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _pnp_odata__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @pnp/odata */ "xT8B");
/* harmony import */ var _webs_types__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../webs/types */ "dVsc");
/* harmony import */ var _types__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./types */ "4tXH");
/* harmony import */ var _utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../utils/escapeQueryStrValue */ "q1Lp");




Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "folders", _types__WEBPACK_IMPORTED_MODULE_2__["Folders"]);
Object(_pnp_odata__WEBPACK_IMPORTED_MODULE_0__["addProp"])(_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"], "rootFolder", _types__WEBPACK_IMPORTED_MODULE_2__["Folder"], "rootFolder");
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.getFolderByServerRelativeUrl = function (folderRelativeUrl) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_2__["Folder"])(this, "getFolderByServerRelativeUrl('" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_3__["escapeQueryStrValue"])(folderRelativeUrl) + "')");
};
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.getFolderByServerRelativePath = function (folderRelativeUrl) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_2__["Folder"])(this, "getFolderByServerRelativePath(decodedUrl='" + Object(_utils_escapeQueryStrValue__WEBPACK_IMPORTED_MODULE_3__["escapeQueryStrValue"])(folderRelativeUrl) + "')");
};
_webs_types__WEBPACK_IMPORTED_MODULE_1__["_Web"].prototype.getFolderById = function (uniqueId) {
    return Object(_types__WEBPACK_IMPORTED_MODULE_2__["Folder"])(this, "getFolderById('" + uniqueId + "')");
};
//# sourceMappingURL=web.js.map

/***/ }),

/***/ "z067":
/*!************************************************!*\
  !*** ./node_modules/@pnp/sp/utils/metadata.js ***!
  \************************************************/
/*! exports provided: metadata */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "metadata", function() { return metadata; });
function metadata(type) {
    return {
        "__metadata": { "type": type },
    };
}
//# sourceMappingURL=metadata.js.map

/***/ }),

/***/ "zYJ9":
/*!*********************************************!*\
  !*** ./node_modules/@pnp/sp/forms/types.js ***!
  \*********************************************/
/*! exports provided: _Forms, Forms, _Form, Form */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Forms", function() { return _Forms; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Forms", function() { return Forms; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "_Form", function() { return _Form; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Form", function() { return Form; });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! tslib */ "LVfT");
/* harmony import */ var _sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../sharepointqueryable */ "E6kk");
/* harmony import */ var _decorators__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ../decorators */ "hMpi");
/* harmony import */ var _telemetry__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ../telemetry */ "h3gW");




/**
 * Describes a collection of Form objects
 *
 */
var _Forms = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Forms, _super);
    function _Forms() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    /**
     * Gets a form by id
     *
     * @param id The guid id of the item to retrieve
     */
    _Forms.prototype.getById = function (id) {
        return _telemetry__WEBPACK_IMPORTED_MODULE_3__["tag"].configure(Form(this).concat("('" + id + "')"), "fos.getById");
    };
    _Forms = Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__decorate"])([
        Object(_decorators__WEBPACK_IMPORTED_MODULE_2__["defaultPath"])("forms")
    ], _Forms);
    return _Forms;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableCollection"]));

var Forms = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Forms);
/**
 * Describes a single of Form instance
 *
 */
var _Form = /** @class */ (function (_super) {
    Object(tslib__WEBPACK_IMPORTED_MODULE_0__["__extends"])(_Form, _super);
    function _Form() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return _Form;
}(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["_SharePointQueryableInstance"]));

var Form = Object(_sharepointqueryable__WEBPACK_IMPORTED_MODULE_1__["spInvokableFactory"])(_Form);
//# sourceMappingURL=types.js.map

/***/ }),

/***/ "zqEK":
/*!*******************************************!*\
  !*** ./node_modules/@pnp/common/index.js ***!
  \*******************************************/
/*! exports provided: objectToMap, mergeMaps, setup, RuntimeConfigImpl, RuntimeConfig, mergeHeaders, mergeOptions, getADALResource, FetchClient, BearerTokenFetchClient, SPFxAdalClient, PnPClientStorageWrapper, PnPClientStorage, getCtxCallback, dateAdd, combine, getRandomString, getGUID, isFunc, objectDefinedNotNull, isArray, assign, isUrlAbsolute, stringIsNullOrEmpty, sanitizeGuid, jsS, hOP, getHashCode, safeGlobal */
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var _collections__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ./collections */ "eVhm");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "objectToMap", function() { return _collections__WEBPACK_IMPORTED_MODULE_0__["objectToMap"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "mergeMaps", function() { return _collections__WEBPACK_IMPORTED_MODULE_0__["mergeMaps"]; });

/* harmony import */ var _libconfig__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./libconfig */ "hyGn");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "setup", function() { return _libconfig__WEBPACK_IMPORTED_MODULE_1__["setup"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RuntimeConfigImpl", function() { return _libconfig__WEBPACK_IMPORTED_MODULE_1__["RuntimeConfigImpl"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "RuntimeConfig", function() { return _libconfig__WEBPACK_IMPORTED_MODULE_1__["RuntimeConfig"]; });

/* harmony import */ var _net__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./net */ "UxNQ");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "mergeHeaders", function() { return _net__WEBPACK_IMPORTED_MODULE_2__["mergeHeaders"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "mergeOptions", function() { return _net__WEBPACK_IMPORTED_MODULE_2__["mergeOptions"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "getADALResource", function() { return _net__WEBPACK_IMPORTED_MODULE_2__["getADALResource"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "FetchClient", function() { return _net__WEBPACK_IMPORTED_MODULE_2__["FetchClient"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "BearerTokenFetchClient", function() { return _net__WEBPACK_IMPORTED_MODULE_2__["BearerTokenFetchClient"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "SPFxAdalClient", function() { return _net__WEBPACK_IMPORTED_MODULE_2__["SPFxAdalClient"]; });

/* harmony import */ var _storage__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./storage */ "9TEK");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PnPClientStorageWrapper", function() { return _storage__WEBPACK_IMPORTED_MODULE_3__["PnPClientStorageWrapper"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "PnPClientStorage", function() { return _storage__WEBPACK_IMPORTED_MODULE_3__["PnPClientStorage"]; });

/* harmony import */ var _util__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./util */ "97IE");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "getCtxCallback", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["getCtxCallback"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "dateAdd", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["dateAdd"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "combine", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["combine"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "getRandomString", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["getRandomString"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "getGUID", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["getGUID"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "isFunc", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["isFunc"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "objectDefinedNotNull", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["objectDefinedNotNull"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "isArray", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["isArray"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "assign", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["assign"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "isUrlAbsolute", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["isUrlAbsolute"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "stringIsNullOrEmpty", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["stringIsNullOrEmpty"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "sanitizeGuid", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["sanitizeGuid"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "jsS", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["jsS"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "hOP", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["hOP"]; });

/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "getHashCode", function() { return _util__WEBPACK_IMPORTED_MODULE_4__["getHashCode"]; });

/* harmony import */ var _safe_global__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./safe-global */ "08ZC");
/* harmony reexport (safe) */ __webpack_require__.d(__webpack_exports__, "safeGlobal", function() { return _safe_global__WEBPACK_IMPORTED_MODULE_5__["safeGlobal"]; });







//# sourceMappingURL=index.js.map

/***/ })

/******/ })});;
//# sourceMappingURL=announcements.bundle.js.map