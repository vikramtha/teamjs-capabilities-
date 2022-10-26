import { useContext, useState } from "react";
import { Image, Menu, Button, Table } from "@fluentui/react-northstar";
import "./Welcome.css";
import { EditCode } from "./EditCode";

import { AzureFunctions } from "./AzureFunctions";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useData } from "@microsoft/teamsfx-react";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import { TeamsFxContext } from "../Context";
import { app, appInstallDialog, executeDeepLink, pages, calendar,call, dialog, location, mail, menus, sharing, barCode, chat, geoLocation, profile,people, search, stageView, teamsCore, video, webStorage} from "@microsoft/teams-js";
import { browserName, CustomView, isMobile } from 'react-device-detect';

const header = {
  items: ['Capability', 'Host Support', 'Notes'],
}
 export const InstallDialog = () => {
 
      return (
          <Button onClick={async () => {
            await appInstallDialog.openAppInstallDialog({
              
              appId: 'com.microsoft.teamspace.tab.youtube'
          });
          }}>
              Open App Install Dialogs
          </Button>
      )
}

export const callButton = () => {
  return(
    <Button onClick={async () => {
      await call.startCall({
        targets: ['vikramtha@microsoft.com'],
        requestedModalities: [call.CallModalities.Audio],
        source: 'source',
    });
}}>
    Make Call
</Button>
)
};
const rows = [
  {
    key: 1,
    items: ['appInstallDialog', appInstallDialog.isSupported().toString(), <InstallDialog/> ],
  },
  {
    key: 2,
    items: ['barCode', barCode.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['Calendar', calendar.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['Call', call.isSupported().toString(), <callButton/>],
  },
  {
    key: 3,
    items: ['Chat', chat.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['geoLocation', geoLocation.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['geoLocation.map', geoLocation.map.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['location', location.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['mail', mail.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['Profile', profile.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['pages.appButton', pages.appButton.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['pages.tabs',pages.tabs.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['pages.backStack', pages.backStack.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['pages', pages.isSupported().toString(), 'None'],
  },

  {
    key: 3,
    items: ['people', people.isSupported().toString(), 'None'],
  },

  

  {
    key: 3,
    items: ['dialog', dialog.isSupported().toString(), 'None'],
  },

  {
    key: 3,
    items: ['dialog.bot', dialog.bot.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['dialog.update', dialog.update.isSupported().toString(), 'None'],
  },{
    key: 3,
    items: ['menus', menus.isSupported().toString(), 'None'],
  },{
    key: 3,
    items: ['search', search.isSupported().toString(), 'None'],
  },{
    key: 3,
    items: ['sharing', sharing.isSupported().toString(), 'None'],
  },{
    key: 3,
    items: ['stageView', stageView.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['teamsCore', teamsCore.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['video', video.isSupported().toString(), 'None'],
  },
  {
    key: 3,
    items: ['webStorage', webStorage.isSupported().toString(), 'None'],
  }

]

export function Welcome(props) {

  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };

  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const steps = ["local", "azure", "publish"];
  const friendlyStepsName = {
    local: "1. Build your app locally",
    azure: "2. Provision and Deploy to the Cloud",
    publish: "3. Publish to Teams",
  };
  const [selectedMenuItem, setSelectedMenuItem] = useState("local");
  const items = steps.map((step) => {
    return {
      key: step,
      content: friendlyStepsName[step] || "",
      onClick: () => setSelectedMenuItem(step),
    };
  });

  const { teamsfx } = useContext(TeamsFxContext);
  const { loading, data, error } = useData(async () => {
    if (teamsfx) {
      const userInfo = await teamsfx.getUserInfo();
      return userInfo;
    }
  });
  const userName = (loading || error) ? "": data.displayName;
  const hubName = useData(async () => {
    await app.initialize();
    const context = await app.getContext();
    return context.app.host.name;
  })?.data;
  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        
        <h1 className="center">Welcome{userName ? ", " + userName : ""}!</h1>
        {hubName && (
          <p className="center">Your current host is: {hubName}</p>
          
        )}
        <div>
        <Table header={header} rows={rows} aria-label="Static table" />
         </div>
         <CustomView condition={browserName === "Chrome"}>
        <div><p className="center">Your current browser is Chrome</p></div>
        </CustomView>
        <CustomView condition={browserName === "Edge"}>
        <div><p className="center">Your current browser is Edge</p></div>

        </CustomView>
        
          <div><p  className="center"> {isMobile? 'This is a mobile device': 'This is a desktop device'} </p></div> 
          
      </div>
      
    </div>
  );
}
