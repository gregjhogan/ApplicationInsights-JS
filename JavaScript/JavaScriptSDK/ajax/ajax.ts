﻿/// <reference path="../logging.ts" />
/// <reference path="../util.ts" />
/// <reference path="./ajaxUtils.ts" />
/// <reference path="./ajaxRecord.ts" />

module Microsoft.ApplicationInsights {
    "use strict";

    export class XMLHttpRequestInstrumented implements XMLHttpRequest {
        msCaching: string;
        onreadystatechange: (ev: ProgressEvent) => any;
        readyState: number;
        response: any;
        responseBody: any;
        responseText: string;
        responseType: string;
        responseXML: any;
        status: number;
        statusText: string;
        timeout: number;
        upload: XMLHttpRequestUpload;
        withCredentials: boolean;
        abort(): void;
        getAllResponseHeaders(): string;
        getResponseHeader(header: string): string;
        msCachingEnabled(): boolean;
        open(method: string, url: string, async?: boolean, user?: string, password?: string): void;
        overrideMimeType(mime: string): void;
        send(data?: Document): void;
        send(data?: string): void;
        send(data?: any): void;
        setRequestHeader(header: string, value: string): void;
        DONE: number;
        HEADERS_RECEIVED: number;
        LOADING: number;
        OPENED: number;
        UNSENT: number;
        addEventListener(type: "abort", listener: (ev: UIEvent) => any, useCapture?: boolean): void;
        addEventListener(type: "error", listener: (ev: ErrorEvent) => any, useCapture?: boolean): void;
        addEventListener(type: "load", listener: (ev: Event) => any, useCapture?: boolean): void;
        addEventListener(type: "loadend", listener: (ev: ProgressEvent) => any, useCapture?: boolean): void;
        addEventListener(type: "loadstart", listener: (ev: Event) => any, useCapture?: boolean): void;
        addEventListener(type: "progress", listener: (ev: ProgressEvent) => any, useCapture?: boolean): void;
        addEventListener(type: "readystatechange", listener: (ev: ProgressEvent) => any, useCapture?: boolean): void;
        addEventListener(type: "timeout", listener: (ev: ProgressEvent) => any, useCapture?: boolean): void;
        addEventListener(type: string, listener: EventListenerOrEventListenerObject, useCapture?: boolean): void;
    }



    //public static onreadyStateChangeCallback() {
    //if (AjaxMonitor.isMonitoredInstance(<XMLHttpRequestInstrumented>this)) {
    //    if (this.onreadystatechange !== onreadystatechangeWrapper) {

    //        if (this.readyState < 3) {

    //            // it is possible to define onreadystatechange event after xhr.send method was invoked.
    //            // intercepting xhr.onreadystatechange in order to measure callback time
    //            interceptOnReadyStateChange.call(this);
    //        }
    //        else {

    //            // On Firefox 13- we cannot override readystatechange, because it is not a function. 
    //            // In this case we don't include callback time in Ajax Total time on this browser
    //            onReadStateChangePrefix.call(this);
    //            onReadyStateChangePostfix.call(this);
    //        }
    //    }
    //}
 
    export declare var XMLHttpRequestInstrumented: {
        prototype: XMLHttpRequest;

        ajaxData: ajaxRecord;
    }

    export class AjaxMonitor {

        private appInsights: Microsoft.ApplicationInsights.AppInsights;
        private static initiated: boolean;

        constructor(appInsights: Microsoft.ApplicationInsights.AppInsights) {
            this.appInsights = appInsights;
            AjaxMonitor.initiated = false;
            this.Init();
        }

        ///<summary>The main function that needs to be called in order to start Ajax Monitoring</summary>
        public Init() {
            if (this.supportMonitoring()) {
                this.interceptOpen();
                this.interceptSetRequestHeader();
                this.interceptSend();
                this.interceptAbort();
                AjaxMonitor.initiated = true;
            }
        }


        ///<summary>Function that returns property name which will identify that monitoring for given instance of XmlHttpRequest is disabled</summary>
        public static GetDisabledPropertyName = function () {
            return "Microsoft_ApplicationInsights_BypassAjaxInstrumentation";
        }

        ///<summary>Verifies that particalar instance of XMLHttpRequest needs to be monitored</summary>
        ///<param name="excludeAjaxDataValidation">Optional parameter. True if ajaxData must be excluded from verification</param>
        ///<returns type="bool">True if instance needs to be monitored, otherwise false</returns>
        public static isMonitoredInstance(instance: XMLHttpRequestInstrumented, excludeAjaxDataValidation) {

            // checking to see that all interested functions on xhr were intercepted
            return AjaxMonitor.initiated

            // checking on ajaxData to see that it was not removed in user code
                && (excludeAjaxDataValidation === true || !extensions.IsNullOrUndefined(instance.ajaxData))

            // check that this instance is not not used by ajax call performed inside client side monitoring to send data to collector
                && instance[AjaxMonitor.GetDisabledPropertyName()] !== true;

        }

        ///<summary>Determines whether ajax monitoring can be enabled on this document</summary>
        ///<returns>True if Ajax monitoring is supported on this page, otherwise false</returns>
        public supportMonitoring() {
            var result = false;
            if (!extensions.IsNullOrUndefined(XMLHttpRequest)) {
                result = true;
            }

            return result;
        }


        public static attachToOnReadyStateChange(instance: XMLHttpRequestInstrumented) {
            instance.ajaxData.onreadystatechangeCallbackAttached = commands.AttachEvent(this, "readystatechange", onreadyStateChangeCallback);
        }

        public interceptOpen() {
            var originalOpen = XMLHttpRequest.prototype.open;
            XMLHttpRequest.prototype.open = function (method, url, async) {
                if (AjaxMonitor.isMonitoredInstance(<XMLHttpRequestInstrumented>this, true)) {
                    this.ajaxData = new ajaxRecord();
                    AjaxMonitor.attachToOnReadyStateChange.call(this);
                    try {
                        this.ajaxData.method = method;
                        this.ajaxData.requestUrl = url;
                        this.ajaxData.requestSize += url.length;

                        // If not set async defaults to true 
                        this.ajaxData.async = extensions.IsNullOrUndefined(async) ? true : async;

                    } catch (e) {
                        // TODO
                    }
                }

                return originalOpen.apply(this, arguments);
            };
        }

        public interceptSend() {
            var originalSend = XMLHttpRequest.prototype.send;
            XMLHttpRequest.prototype.send = function (content) {
                AjaxMonitor.sendPrefixInterceptor(this, content);
                return originalSend.apply(this, arguments);
            };
        }

        public static sendPrefixInterceptor(instance: XMLHttpRequest, content) {
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
                } catch (e) {
                    // TODO
                }
            }
        }

        public interceptAbort() {
            var originalAbort = XMLHttpRequest.prototype.abort;
            XMLHttpRequest.prototype.abort = function () {
                if (this.isMonitoredInstance.call(this)) {
                    this.ajaxData.aborted = 1;
                }

                return originalAbort.apply(this, arguments);
            };
        }

        ///<summary>Intercept onreadystatechange callback</summary>
        ///<returns>True, if onreadystatechange is intercepted, otherwise false</returns>
        public interceptOnReadyStateChange() {
            var result = false;

            // do not intercept onreadystatechange if it is defined and not a function, because we are not able to call original function in this case, which happends on Firefox 13 and lower
            if (extensions.IsNullOrUndefined(this.onreadystatechange) || (typeof (this.onreadystatechange) === "function")) {
                this.ajaxData.originalOnreadystatechage = this.onreadystatechange;
                this.onreadystatechange = onreadystatechangeWrapper;
                result = true;
            }

            return result;
        }

        public onreadystatechangeWrapper() {
            if (isMonitoredInstance.call(this)) {
                onReadStateChangePrefix.call(this);
                try {

                    // customer's callback can raise exception. We need to proceed monitor ajax call in this case as well.
                    if (!extensions.IsNullOrUndefined(this.ajaxData.originalOnreadystatechage)) {
                        this.ajaxData.originalOnreadystatechage.call(this);
                    }
                } catch (ex) {
                    this.ajaxData.clientFailure = 1;
                    throw ex;

                } finally {
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

        public onReadStateChangePrefix() {
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
        }

        public onReadyStateChangePostfix() {
            if (this.readyState === 4) {
                onAjaxComplete.call(this);
            }
        }

        public onAjaxComplete() {
            commands.TryCatchTraceWrapper.call(this, "publishData", function () {
                this.ajaxData.CalculateMetrics();

                var successStatuses = [200, 201, 202, 203, 204, 301, 302, 303, 304];

                this.appInsights.trackAjax(
                    this.ajaxData.getAbsoluteUrl(),
                    this.ajaxData.async,
                    this.ajaxData.ajaxTotalDuration,
                    successStatuses.indexOf(+this.ajaxData.status) != -1
                    );
            });

            commands.TryCatchTraceWrapper.call(this, "deleteAjaxData", function () {
                commands.DetachEvent(this, "readystatechange", onreadyStateChangeCallback);
                delete this.ajaxData;
            });
        }

        public collectResponseData() {
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
            } catch (e) {
                _InternalLogging.throwInternalNonUserActionable(
                    LoggingSeverity.CRITICAL,
                    "Failed to collect response data: "
                    + Microsoft.ApplicationInsights.Util.dump(e));
            }
        }

        public interceptSetRequestHeader() {
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
        }
    }
}