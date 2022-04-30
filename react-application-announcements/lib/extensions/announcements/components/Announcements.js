import * as React from 'react';
import { Link, MessageBar, MessageBarType } from '@microsoft/office-ui-fabric-react-bundle';
import { Web } from "@pnp/sp/presets/all";
import { useEffect, useState } from 'react';
import * as strings from 'announcementsStrings';
import { QUALIFIED_NAME } from '../AnnouncementsApplicationCustomizer';
export default function RenderAnnouncements(props) {
    // Two local state variables with their setter
    var _a = useState([]), announcements = _a[0], setAnnouncements = _a[1];
    var _b = useState([]), acknowledgedAnnouncements = _b[0], setAcknowledgedAnnouncements = _b[1];
    // Use an effect to query the list data only once,
    // not on every render. The effect will be re-run if
    // props.siteUrl or props.listName changes
    useEffect(function () {
        if (window.localStorage) {
            var items = window.localStorage.getItem(QUALIFIED_NAME);
            if (items) {
                setAcknowledgedAnnouncements(JSON.parse(items));
            }
        }
        // Use PnP JS to query SharePoint
        var now = new Date().toISOString();
        Web(props.siteUrl)
            .lists.getByTitle(props.listName)
            .items
            .filter("(Locale eq '" + props.culture + "' or Locale eq null) and (StartDateTime le datetime'" + now + "' or StartDateTime eq null) and (EndDateTime ge datetime'" + now + "' or EndDateTime eq null)")
            .select("ID", "Title", "Announcement", "Urgent", "Link", "Locale", "StartDateTime", "EndDateTime")
            .get()
            .then(setAnnouncements);
    }, [props.siteUrl, props.listName]);
    var announcementElements = announcements
        .filter(function (announcement) { return acknowledgedAnnouncements.indexOf(announcement.ID) < 0; })
        .map(function (announcement) { return React.createElement(MessageBar, { messageBarType: (announcement.Urgent ? MessageBarType.error : MessageBarType.warning), isMultiline: false, onDismiss: function () {
            // On dismiss, add the current announcement id to the array 
            // STORAGE_KEY item in localStorage so it is remembered locally
            var items = JSON.parse(window.localStorage.getItem(QUALIFIED_NAME)) || [];
            items.push(announcement.ID);
            window.localStorage.setItem(QUALIFIED_NAME, JSON.stringify(items));
            setAcknowledgedAnnouncements(items);
        }, dismissButtonAriaLabel: strings.Close },
        React.createElement("strong", null, announcement.Title),
        "\u00A0",
        React.createElement("span", { dangerouslySetInnerHTML: { __html: announcement.Announcement } }),
        announcement.Link && React.createElement(Link, { href: announcement.Link.Url, target: "_blank" }, announcement.Link.Description)); });
    return React.createElement("div", null, announcementElements);
}
//# sourceMappingURL=Announcements.js.map