import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
import "jquery";
import 'alertifyjs';
import '../../ExternalRef/css/style.css';
import '../../ExternalRef/css/alertify.min.css';
import '../../ExternalRef/css/bootstrap-datepicker.min.css';
export interface IProjectTrackerWebPartProps {
    ListName: string;
}
export default class ProjectTrackerWebPart extends BaseClientSideWebPart<IProjectTrackerWebPartProps> {
    onInit(): Promise<void>;
    render(): void;
    mandatoryValidation(): void;
    fetchStatus(): void;
    fetchEIMTicketNo(): void;
    fetchTasktype(): void;
    fetchEscReason(): void;
    fetchNotificationAction(): void;
    saveData(): void;
    protected readonly disableReactivePropertyChanges: boolean;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
