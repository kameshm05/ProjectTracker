var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from "@microsoft/sp-loader";
import "jquery";
require("bootstrap");
import 'alertifyjs';
// import "jquery-ui";
import '../../ExternalRef/css/style.css';
import '../../ExternalRef/css/alertify.min.css';
import '../../ExternalRef/css/bootstrap-datepicker.min.css';
require('../../ExternalRef/js/bootstrap-datepicker.min.js');
var alertify = require('../../ExternalRef/js/alertify.min.js');
import { sp } from "@pnp/sp";
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
import * as strings from 'ProjectTrackerWebPartStrings';
var listname, startTime, completionTime;
var isAllFilled = true;
var ProjectTrackerWebPart = (function (_super) {
    __extends(ProjectTrackerWebPart, _super);
    function ProjectTrackerWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ProjectTrackerWebPart.prototype.onInit = function () {
        var _this = this;
        return _super.prototype.onInit.call(this).then(function (_) {
            sp.setup({
                spfxContext: _this.context
            });
        });
    };
    ProjectTrackerWebPart.prototype.render = function () {
        var _this = this;
        listname = this.properties.ListName;
        this.domElement.innerHTML = "<div id='form-start' class=\"row\"><div class=\"col-sm-8\"><div class=\"form-group forms-required \"><label  id='lblSiteName' class='label-bold'>1.Site Common Name</label><input class=\"form-control\" type=\"text\" id=\"txtSiteCommonName\"></div><div class=\"form-group forms-required lblStatusfull\"><label id='lblStatus' class='label-bold'>2.Status</label></div><div class=\"form-group forms-required lblTaskTypefull\"><label id='lblTaskType' class='label-bold'>3.Task Type</label></div><div class=\"form-group forms-required\"><label id='lblDateAssigned' class='label-bold'>4.Date Assigned</label><br><input type=\"text\" id=\"txtDate\" class=\"form-control form-control-datepicker\"></div><div class=\"form-group\"><label id='lblIWOSTicket' class='label-bold'>5.IWOS Ticket</label><br><input type=\"text\" id=\"txtIWOSTicket\" class=\"form-control\"></div><div class=\"form-group\"><label id='lblNestTicketURL' class='label-bold'>6.Nest Ticket URL</label><br><input type=\"text\" id=\"txtNestTicketURL\" class=\"form-control\"></div><div class=\"form-group forms-required\"><label id='lblServiceImpacting' class='label-bold'>7.Service Impacting</label></div><div class='form-group'><input type=\"radio\" class=\"radio-stylish\" name=\"ServiceImpacting\" value=\"Yes\" id=\"ServiceYes\"><span class=\"radio-element\"></span><label class=\"stylish-label\" for=\"ServiceYes\">Yes</label></div><div class='form-group'><input type=\"radio\" class=\"radio-stylish\" name=\"ServiceImpacting\" id=\"ServiceNo\" value=\"No\"><span class=\"radio-element\"></span><label class=\"stylish-label\" for=\"ServiceNo\">No</label></div><label id='lblEIMTicketNo' class='label-bold'>8.EIM Ticket No</label><label id='lblEIMTicket' class=\"label-bold\">9.EIM Ticket</label><div class=\"form-group\"><input type=\"text\" id=\"txtEIMTicket\" class=\"form-control\"></div><label id='lblTaskNotes' class='label-bold'>10.Task Notes</label><div class=\"form-group\"><textarea id='txtTaskNotes'  class=\"form-control\"></textarea></div><label id='lblEsclationNeeded' class='label-bold'>11.Escalation Needed</label><div class='form-group'><input type=\"radio\" name=\"EsclationNeeded\"  class=\"radio-stylish\" value=\"Yes\" id=\"EscYes\"><span class=\"radio-element\"></span><label class=\"stylish-label\" for=\"EscYes\">Yes</label></div><div class=\"form-group\"><input type=\"radio\" name=\"EsclationNeeded\" class=\"radio-stylish\" value=\"No\"  id=\"EscNo\"><span class=\"radio-element\"></span><label class=\"stylish-label\" for=\"EscNo\">No</label></div><label id='lblEscalationReason' class='label-bold'>12.Escalation Reason</label><label id='lblNotificationAction' class='label-bold'>13.Notification Action</label><div class=\"form-group\"><input type=\"button\" id=\"btnSave\" class=\"btn btn-primary\" value=\"Save\"></div></div></div>";
        startTime = new Date().toLocaleString();
        $("#txtDate").datepicker();
        this.fetchStatus();
        this.fetchEIMTicketNo();
        this.fetchTasktype();
        this.fetchEscReason();
        this.fetchNotificationAction();
        $('#btnSave').click(function () {
            _this.mandatoryValidation();
        });
    };
    ProjectTrackerWebPart.prototype.mandatoryValidation = function () {
        isAllFilled = true;
        var selectedTasktype = [];
        $("input[name='Tasktype']:checked").each(function () {
            selectedTasktype.push($(this).val());
        });
        if (!listname) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error('Please enter list name');
            isAllFilled = false;
        }
        else if (!$("#txtSiteCommonName").val()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error('Please enter Site Common Name');
            isAllFilled = false;
        }
        else if (!$("input[name='Status']:checked").val()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error('Please select Status');
            isAllFilled = false;
        }
        else if (!selectedTasktype.length) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error('Please select Task Type');
            isAllFilled = false;
        }
        else if (!$("#txtDate").val()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error('Please select Date Assigned');
            isAllFilled = false;
        }
        else if (!$("input[name='ServiceImpacting']:checked").val()) {
            alertify.set('notifier', 'position', 'top-right');
            alertify.error('Please select Service Impacting');
            isAllFilled = false;
        }
        if (isAllFilled) {
            this.saveData();
        }
    };
    ProjectTrackerWebPart.prototype.fetchStatus = function () {
        var getStatusRadiovalues;
        var renderStatusRadioOption = "";
        getStatusRadiovalues = sp.web.lists
            .getByTitle("Status")
            .items.get()
            .then(function (items) {
            for (var i = 0; i < items.length; i++) {
                renderStatusRadioOption += "<div class='form-group'><input class='radio-stylish' type=radio name='Status' value='" + items[i].Status + "' id='radiostatus" + i + "'><span class='radio-element'></span><label class='stylish-label' for='radiostatus" + i + "'>" + items[i].Status + "</label></div>";
            }
            //renderStatusRadioOption+="<br>";
            $(".lblStatusfull").after(renderStatusRadioOption);
        });
    };
    ProjectTrackerWebPart.prototype.fetchEIMTicketNo = function () {
        var getEIMTicketNoRadiovalues;
        var renderEIMTicketNoRadioOption = "";
        getEIMTicketNoRadiovalues = sp.web.lists
            .getByTitle("EIMTicketNo")
            .items.get()
            .then(function (items) {
            for (var i = 0; i < items.length; i++) {
                renderEIMTicketNoRadioOption += "<div class='form-group'><input type=radio class='radio-stylish' name='EIMTicketNo' value='" + items[i].EIMTicketNo + "' id='radioEMINo" + i + "'><span class='radio-element'></span><label class='stylish-label' for='radioEMINo" + i + "'>" + items[i].EIMTicketNo + "</label></div>";
            }
            // renderEIMTicketNoRadioOption+="<br>";
            $("#lblEIMTicketNo").after(renderEIMTicketNoRadioOption);
        });
    };
    ProjectTrackerWebPart.prototype.fetchTasktype = function () {
        var getTaskTypeRadiovalues;
        var renderTasktypeRadioOption = "";
        getTaskTypeRadiovalues = sp.web.lists
            .getByTitle("TaskType")
            .items.get()
            .then(function (items) {
            for (var i = 0; i < items.length; i++) {
                renderTasktypeRadioOption += "<div class='form-group'><input type=checkbox class='radio-stylish' name='Tasktype' value='" + items[i].TaskType + "' id='checkTask" + i + "'><span class='checkbox-element'></span><label class='stylish-label' for='checkTask" + i + "'>" + items[i].TaskType + "</label></div>";
            }
            //renderTasktypeRadioOption+="<br>";
            $(".lblTaskTypefull").after(renderTasktypeRadioOption);
        });
    };
    ProjectTrackerWebPart.prototype.fetchEscReason = function () {
        var getEscReasonRadiovalues;
        var renderEscReasonRadioOption = "";
        getEscReasonRadiovalues = sp.web.lists
            .getByTitle("EscalationReason")
            .items.get()
            .then(function (items) {
            for (var i = 0; i < items.length; i++) {
                renderEscReasonRadioOption += "<div class='form-group'><input type=checkbox class='radio-stylish' name='EscReason' value='" + items[i].EscalationReason + "' id='checkescreason" + i + "'><span class='checkbox-element'></span><label class='stylish-label' for='checkescreason" + i + "'>" + items[i].EscalationReason + "</label></div>";
            }
            //renderEscReasonRadioOption+="<br>";
            $("#lblEscalationReason").after(renderEscReasonRadioOption);
        });
    };
    ProjectTrackerWebPart.prototype.fetchNotificationAction = function () {
        var getNotificationActionRadiovalues;
        var renderNotificationActionRadioOption = "";
        getNotificationActionRadiovalues = sp.web.lists
            .getByTitle("NotificationAction")
            .items.get()
            .then(function (items) {
            for (var i = 0; i < items.length; i++) {
                renderNotificationActionRadioOption += "<div class='form-group'><input type=radio  class='radio-stylish'  name='NotificationAction' value='" + items[i].NotificationAction + "' id='radioNotify" + i + "'><span class='radio-element'></span><label class='stylish-label' for='radioNotify" + i + "'>" + items[i].NotificationAction + "</label></div>";
            }
            //renderNotificationActionRadioOption+="<br>";
            $("#lblNotificationAction").after(renderNotificationActionRadioOption);
        });
    };
    ProjectTrackerWebPart.prototype.saveData = function () {
        completionTime = new Date().toLocaleString();
        var fullName = this.context.pageContext.user.displayName;
        var splitname = fullName.split(' ');
        var selectedTasktype = [];
        $("input[name='Tasktype']:checked").each(function () {
            selectedTasktype.push($(this).val());
        });
        if (selectedTasktype.length > 0)
            var finalselectedTasktype = selectedTasktype.join(';');
        else
            var finalselectedTasktype = "";
        var selectedEscReason = [];
        $("input[name='EscReason']:checked").each(function () {
            selectedEscReason.push($(this).val());
        });
        if (selectedEscReason.length > 0)
            var finalselectedEscReason = selectedEscReason.join(';');
        else
            var finalselectedEscReason = "";
        var objData = {
            StartTime: startTime,
            CompletionTime: completionTime,
            Email: this.context.pageContext.user.email,
            Name: fullName,
            FirstName: splitname[0],
            LastName: splitname[1],
            Status: $("input[name='Status']:checked").val(),
            TaskType: finalselectedTasktype,
            DateAssigned: new Date($('#txtDate').val()).toLocaleDateString(),
            IWOSTicket: $('#txtIWOSTicket').val(),
            NestTicketURL: $('#txtNestTicketURL').val(),
            EIMTicketNo: $("input[name='EIMTicketNo']:checked").val(),
            EIMTicket: $('#txtEIMTicket').val(),
            TaskNotes: $('#txtTaskNotes').val(),
            ServiceImpacting: $("input[name='ServiceImpacting']:checked").val(),
            EscalationNeeded: $("input[name='EsclationNeeded']:checked").val(),
            EscalationReason: finalselectedEscReason,
            NotificationAction: $("input[name='NotificationAction']:checked").val(),
            SiteCommonName: $('#txtSiteCommonName').val()
        };
        var addData = sp.web.lists.getByTitle(listname).items.add({
            StartTime: startTime,
            CompletionTime: completionTime,
            Email: this.context.pageContext.user.email,
            Name: fullName,
            FirstName: splitname[0],
            LastName: splitname[1],
            Status: $("input[name='Status']:checked").val(),
            TaskType: finalselectedTasktype,
            DateAssigned: new Date($('#txtDate').val()).toLocaleDateString(),
            IWOSTicket: $('#txtIWOSTicket').val(),
            NestTicketURL: $('#txtNestTicketURL').val(),
            EIMTicketNo: $("input[name='EIMTicketNo']:checked").val(),
            EIMTicket: $('#txtEIMTicket').val(),
            TaskNotes: $('#txtTaskNotes').val(),
            ServiceImpacting: $("input[name='ServiceImpacting']:checked").val(),
            EscalationNeeded: $("input[name='EsclationNeeded']:checked").val(),
            EscalationReason: finalselectedEscReason,
            NotificationAction: $("input[name='NotificationAction']:checked").val(),
            SiteCommonName: $('#txtSiteCommonName').val()
        }).then(function (iar) {
            alertify.alert("Data added successfully", function () {
                location.reload();
            }).setHeader('Success').set('closable', false);
        }).catch(function (e) {
            alertify.alert("Please provide correct list name", function () {
                location.reload();
            }).setHeader('Warning').set('closable', false);
        });
        ;
    };
    Object.defineProperty(ProjectTrackerWebPart.prototype, "disableReactivePropertyChanges", {
        get: function () {
            return true;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(ProjectTrackerWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    ProjectTrackerWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    groups: [
                        {
                            groupFields: [
                                PropertyPaneTextField("ListName", {
                                    label: strings.listnameFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return ProjectTrackerWebPart;
}(BaseClientSideWebPart));
export default ProjectTrackerWebPart;
//# sourceMappingURL=ProjectTrackerWebPart.js.map