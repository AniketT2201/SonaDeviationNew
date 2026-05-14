import * as React from 'react';
import { useHistory, useLocation } from 'react-router-dom';
import '../../components/Pages/Css/Sidebar.scss';
import { ISonaDeviationProps } from '../ISonaDeviationProps';
import '@fortawesome/fontawesome-free/css/all.min.css';
import logo from "../../assets/SonaPNGLogo.png";
import { sp } from '@pnp/sp';

type LocationState = { from?: string };

const Sidebar = (props: ISonaDeviationProps) => {
  const history = useHistory();
  const location = useLocation();
  const [selectedTabUrl, setSelectedTabUrl] = React.useState<string>("")
  const [tabs, setTabs] = React.useState<Array<{ id: number; title: string; seq: number; url: string }>>([]);
  // const [loading, setLoading] = React.useState(true);
  const username = props.userDisplayName;

  // Setup PnPJS
  React.useEffect(() => {
    sp.setup({ spfxContext: props.context });
  }, []);

  /** ---------------------------------------------------------
   * Load tabs that the user has access to via SharePoint groups
   * --------------------------------------------------------- */

  const loadTabsWithAccess = async () => {
    try {
      // Get current user's SharePoint groups
      const userGroups = await sp.web.currentUser.groups();
      const userGroupIds = userGroups.map(g => g.Id);

      // Get all tabs from "Tabing" list
      const items: any[] = await sp.web.lists
        .getByTitle("DeviationTabbing")
        .items
        .select(
          "Id",
          "Title",
          "SeqNo",
          "PageUrl",
          "TabingViewGroup/Id",
          "TabingViewGroup/Title"
        )
        .expand("TabingViewGroup")();

      // Filter tabs
      const allowedTabs = items
        .filter(tab => {
          const groups = tab.TabingViewGroup || [];

          // ✅ Public tab → no group assigned
          if (groups.length === 0) return true;

          // ✅ User belongs to group → show
          return groups.some((g: any) => userGroupIds.includes(g.Id));
        })
        .map(tab => ({
          id: tab.Id,
          title: tab.Title,
          seq: tab.SeqNo || 999,
          url: tab.PageUrl ? tab.PageUrl.replace(/\s+/g, "") : ""
        }))
        .sort((a, b) => a.seq - b.seq);

      setTabs(allowedTabs);

      if (allowedTabs.length > 0) {
        setSelectedTabUrl(allowedTabs[0].url);
      }
    } catch (err) {
      console.error("Load Tabs Error:", err);
    }
  };

  const getActiveClass = (tabUrl: string) => {
    // current route from browser hash
    const currentRoute = window.location.hash.replace("#", "") || "/";

    // route from tab url
    const tabRoute = tabUrl.split("#")[1] || "/";

    return currentRoute === tabRoute ? "active" : "";
  };


  /** ---------------------------------------------------------
   * Navigate to page
   * --------------------------------------------------------- */

  React.useEffect(() => {
    loadTabsWithAccess();
  }, []);

  return (
    <div className="sidebar">
      <div className="sidehead">
        <div className="logo">
          <img src={logo} width="25px" height="25px" />
        </div>
        <div className="sidehead-right">SONA COMSTAR</div>
      </div>

      <div className="sidehead-user">
        <i className="fas fa-user" style={{ marginLeft: "20px", marginRight: "15px" }}></i>
        {username}
      </div>

      <ul className="nav">
        {tabs.map(tab => (
          <li className="nav-item" key={tab.id}>
            <a
              href={tab.url}
              className={`nav-link ${getActiveClass(tab.url)}`}
              style={{ cursor: "pointer" }}
            >
              {tab.title}
            </a>

          </li>
        ))}

      </ul>
    </div>
  );
};

export default Sidebar;