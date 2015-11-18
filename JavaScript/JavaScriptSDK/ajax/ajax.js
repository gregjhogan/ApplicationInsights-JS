/// <reference path="../logging.ts" />
/// <reference path="../util.ts" />
/// <reference path="./ajaxUtils.ts" />
/// <reference path="./ajaxRecord.ts" />
define(["require", "exports"], function (require, exports) {
    var Microsoft;
    (function (Microsoft) {
        var ApplicationInsights;
        (function (ApplicationInsights) {
            "use strict";
            { }
            ;
        })(ApplicationInsights = Microsoft.ApplicationInsights || (Microsoft.ApplicationInsights = {}));
    })(Microsoft || (Microsoft = {}));
    var AjaxMonitor = (function () {
        function AjaxMonitor(appInsights) {
            this.appInsights = appInsights;
            AjaxMonitor.initiated = false;
            this.Init();
        }
        ///<summary>The main function that needs to be called in order to start Ajax Monitoring</summary>
        AjaxMonitor.prototype.Init = function () {
            if (this.supportMonitoring()) {
                this.interceptOpen();
                this.interceptSetRequestHeader();
                this.interceptSend();
                this.interceptAbort();
                AjaxMonitor.initiated = true;
            }
        };
        ///<summary>Verifies that particalar instance of XMLHttpRequest needs to be monitored</summary>
        ///<param name="excludeAjaxDataValidation">Optional parameter. True if ajaxData must be excluded from verification</param>
        ///<returns type="bool">True if instance needs to be monitored, otherwise false</returns>
        AjaxMonitor.isMonitoredInstance = function (instance, excludeAjaxDataValidation) {
            // checking to see that all interested functions on xhr were intercepted
            return AjaxMonitor.initiated
                && (excludeAjaxDataValidation === true || !extensions.IsNullOrUndefined(instance.ajaxData))
                && instance[AjaxMonitor.GetDisabledPropertyName()] !== true;
        };
        ///<summary>Determines whether ajax monitoring can be enabled on this document</summary>
        ///<returns>True if Ajax monitoring is supported on this page, otherwise false</returns>
        AjaxMonitor.prototype.supportMonitoring = function () {
            var result = false;
            if (!extensions.IsNullOrUndefined(XMLHttpRequest)) {
                result = true;
            }
            return result;
        };
        AjaxMonitor.attachToOnReadyStateChange = function (instance) {
            instance.ajaxData.onreadystatechangeCallbackAttached = commands.AttachEvent(this, "readystatechange", onreadyStateChangeCallback);
        };
        AjaxMonitor.prototype.interceptOpen = function () {
            var originalOpen = XMLHttpRequest.prototype.open;
            XMLHttpRequest.prototype.open = function (method, url, async) {
                if (AjaxMonitor.isMonitoredInstance(this, true)) {
                    this.ajaxData = new ajaxRecord();
                    AjaxMonitor.attachToOnReadyStateChange.call(this);
                    try {
                        this.ajaxData.method = method;
                        this.ajaxData.requestUrl = url;
                        this.ajaxData.requestSize += url.length;
                        // If not set async defaults to true 
                        this.ajaxData.async = extensions.IsNullOrUndefined(async) ? true : async;
                    }
                    catch (e) {
                    }
                }
                return originalOpen.apply(this, arguments);
            };
        };
        AjaxMonitor.prototype.interceptSend = function () {
            var originalSend = XMLHttpRequest.prototype.send;
            XMLHttpRequest.prototype.send = function (content) {
                AjaxMonitor.sendPrefixInterceptor(this, content);
                return originalSend.apply(this, arguments);
            };
        };
        AjaxMonitor.sendPrefixInterceptor = function (instance, content) {
            if (this.isMonitoredInstance.call(this)) {
                try {
                    if (!extensions.IsNullOrUndefined(content) && !extensions.IsNullOrUndefined(content.length)) {
                        // http://www.w3.org/TR/XMLHttpRequest/: If the request method is a case-sensitive match for GET or HEAD act as if data is null.
                        if (instance.ajaxData.method !== "GET" && instance.ajaxData.method !== "HEAD") {
                            instance.ajaxData.requestSize += content.length;
                        }
                    }
                    this.ajaxData.requestSentTime = dateTime.Now();
                    this.ajaxData.loadingRequest = document.readyState === "loading";
                    if (!this.ajaxData.onreadystatechangeCallbackAttached) {
                        // IE 8 and below does not support xmlh.addEventListener. This the last place for the browsers that does not support addEventListenener to intercept onreadystatechange
                        var that = this;
                        setTimeout(function () {
                            if (that.readyState === 4) {
                                // ajax is cached, onreadystatechange didn't fire, but it is completed
                                commands.TryCatchTraceWrapper.call(that, "readyState(4)", this.collectResponseData);
                                this.onAjaxComplete.call(that);
                            }
                            else {
                                this.interceptOnReadyStateChange.call(that);
                            }
                        }, 5);
                    }
                }
                catch (e) {
                }
            }
        };
        AjaxMonitor.prototype.interceptAbort = function () {
            var originalAbort = XMLHttpRequest.prototype.abort;
            XMLHttpRequest.prototype.abort = function () {
                if (this.isMonitoredInstance.call(this)) {
                    this.ajaxData.aborted = 1;
                }
                return originalAbort.apply(this, arguments);
            };
        };
        ///<summary>Intercept onreadystatechange callback</summary>
        ///<returns>True, if onreadystatechange is intercepted, otherwise false</returns>
        AjaxMonitor.prototype.interceptOnReadyStateChange = function () {
            var result = false;
            // do not intercept onreadystatechange if it is defined and not a function, because we are not able to call original function in this case, which happends on Firefox 13 and lower
            if (extensions.IsNullOrUndefined(this.onreadystatechange) || (typeof (this.onreadystatechange) === "function")) {
                this.ajaxData.originalOnreadystatechage = this.onreadystatechange;
                this.onreadystatechange = onreadystatechangeWrapper;
                result = true;
            }
            return result;
        };
        AjaxMonitor.prototype.onreadystatechangeWrapper = function () {
            if (isMonitoredInstance.call(this)) {
                onReadStateChangePrefix.call(this);
                try {
                    // customer's callback can raise exception. We need to proceed monitor ajax call in this case as well.
                    if (!extensions.IsNullOrUndefined(this.ajaxData.originalOnreadystatechage)) {
                        this.ajaxData.originalOnreadystatechage.call(this);
                    }
                }
                catch (ex) {
                    this.ajaxData.clientFailure = 1;
                    throw ex;
                }
                finally {
                    if (!extensions.IsNullOrUndefined(this.ajaxData.originalOnreadystatechage)) {
                        commands.TryCatchTraceWrapper.call(this, "callbackFinishedTime", function () {
                            if (this.readyState === 4) {
                                this.ajaxData.callbackFinishedTime = dateTime.Now();
                            }
                        });
                    }
                    onReadyStateChangePostfix.call(this);
                }
            }
        };
        ;
        AjaxMonitor.prototype.onReadStateChangePrefix = function () {
            switch (this.readyState) {
                case 3:
                    commands.TryCatchTraceWrapper.call(this, "readyState(3)", function () {
                        this.ajaxData.responseStartedTime = dateTime.Now();
                    });
                    break;
                case 4:
                    commands.TryCatchTraceWrapper.call(this, "readyState(4)", collectResponseData);
                    break;
            }
        };
        AjaxMonitor.prototype.onReadyStateChangePostfix = function () {
            if (this.readyState === 4) {
                onAjaxComplete.call(this);
            }
        };
        AjaxMonitor.prototype.onAjaxComplete = function () {
            commands.TryCatchTraceWrapper.call(this, "publishData", function () {
                this.ajaxData.CalculateMetrics();
                var successStatuses = [200, 201, 202, 203, 204, 301, 302, 303, 304];
                this.appInsights.trackAjax(this.ajaxData.getAbsoluteUrl(), this.ajaxData.async, this.ajaxData.ajaxTotalDuration, successStatuses.indexOf(+this.ajaxData.status) != -1);
            });
            commands.TryCatchTraceWrapper.call(this, "deleteAjaxData", function () {
                commands.DetachEvent(this, "readystatechange", onreadyStateChangeCallback);
                delete this.ajaxData;
            });
        };
        AjaxMonitor.prototype.collectResponseData = function () {
            var currentTime = dateTime.Now();
            var self = this;
            self.ajaxData.responseFinishedTime = currentTime;
            // Next condition is TRUE sometimes, when ajax request is not authorised by server.
            // See TFS #11632 for details.
            if (self.ajaxData.responseStartedTime === null) {
                self.ajaxData.responseStartedTime = currentTime;
            }
            // FF throws exception on accessing properties of xhr when network error occured during ajax call
            // http://helpful.knobs-dials.com/index.php/Component_returned_failure_code:_0x80040111_(NS_ERROR_NOT_AVAILABLE)
            try {
                this.ajaxData.status = this.status;
                this.ajaxData.contentType = this.getResponseHeader("Content-Type");
                this.ajaxData.responseSize = this.responseText.length;
                this.ajaxData.responseSize += this.getAllResponseHeaders().length;
                //add 'HTTP/1.1 200 OK' length
                this.ajaxData.responseSize += 17;
            }
            catch (e) {
                _InternalLogging.throwInternalNonUserActionable(LoggingSeverity.CRITICAL, "Failed to collect response data: "
                    + Microsoft.ApplicationInsights.Util.dump(e));
            }
        };
        AjaxMonitor.prototype.interceptSetRequestHeader = function () {
            var originalSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
            XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
                if (this.isMonitoredInstance.call(this)) {
                    commands.TryCatchTraceWrapper.call(this, "Adding size of header to total request size", function () {
                        // 2 is the length of ": " which is added to each header
                        this.ajaxData.requestSize += stringUtils.GetLength(name) + stringUtils.GetLength(value) + 2;
                    });
                }
                return originalSetRequestHeader.apply(this, arguments);
            };
        };
        ///<summary>Function that returns property name which will identify that monitoring for given instance of XmlHttpRequest is disabled</summary>
        AjaxMonitor.GetDisabledPropertyName = function () {
            return "Microsoft_ApplicationInsights_BypassAjaxInstrumentation";
        };
        return AjaxMonitor;
    })();
    exports.AjaxMonitor = AjaxMonitor;
});
//# sourceMappingURL=ajax.js.map