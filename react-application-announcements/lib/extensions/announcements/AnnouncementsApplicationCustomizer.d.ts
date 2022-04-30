import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
export declare const QUALIFIED_NAME = "Extension.ApplicationCustomizer.Announcements";
export interface IAnnouncementsApplicationCustomizerProperties {
    siteUrl: string;
    listName: string;
}
export default class AnnouncementsApplicationCustomizer extends BaseApplicationCustomizer<IAnnouncementsApplicationCustomizerProperties> {
    protected onInit(): Promise<void>;
}
//# sourceMappingURL=AnnouncementsApplicationCustomizer.d.ts.map