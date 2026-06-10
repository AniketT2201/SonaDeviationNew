import * as React from 'react';
import { HashRouter as Router, Switch, Route, useLocation } from 'react-router-dom';
import type { ISonaDeviationProps } from './ISonaDeviationProps';
import Sidebar from '../components/Pages/Sidebar';
import { NewRequest } from '../components/Pages/DeviationRequest';
import { DeviationViewForm } from '../components/Pages/DeviationViewform';
import { DeviationEditForm} from '../components/Pages/DeviationEditForm';
import { DeviationApproverForm } from '../components/Pages/DeviationApprovalform';
import { InitiatorDashboard } from '../components/Pages/InitiatorDashboard';
import { COODashboard } from '../components/Pages/COODashboard';
import { QualityDashboard } from '../components/Pages/QualityDashboard';
import { EngineeringDashboard } from '../components/Pages/EngineeringDashboard';
import { PlantHeadDashboard } from '../components/Pages/PlantHeadDashboard';
import { ReportDashboard } from '../components/Pages/ReportDashboard';
import Reports from './Pages/Reports';

// 👇 Separate Layout Component
const Layout: React.FC<ISonaDeviationProps> = (props) => {
  const location = useLocation(); // ✅ Now inside Router

  const hideSidebar =
    location.pathname.startsWith("/NewRequest") ||
    location.pathname.startsWith("/DeviationViewForm") ||
    location.pathname.startsWith("/DeviationEditForm") ||
    location.pathname.startsWith("/DeviationApproverForm");
  return (
    <div className="container-fluid" style={{ display: 'flex', width: '100%' }}>

      {!hideSidebar && <Sidebar {...props} />}

      <div className="main" style={{
        width: hideSidebar ? "100%" : "calc(100% - 250px)",
        transition: "width 0.3s ease"
      }}>
        <Switch>
          <Route exact path="/" render={() => <InitiatorDashboard {...props} />} />
          <Route exact path="/NewRequest" render={() => <NewRequest {...props} />} />
          <Route exact path="/DeviationViewForm/:id" render={() => <DeviationViewForm {...props} />} />
          <Route exact path="/DeviationEditForm/:id" render={() => <DeviationEditForm {...props} />} />
          <Route exact path="/DeviationApproverForm/:id" render={() => <DeviationApproverForm {...props} />} />
          <Route exact path="/QualityDashboard" render={() => <QualityDashboard {...props} />} /> 
          <Route exact path="/COODashboard" render={() => <COODashboard {...props} />} /> 
          <Route exact path="/EngineeringDashboard" render={() => <EngineeringDashboard {...props} />} /> 
          <Route exact path="/PlantHeadDashboard" render={() => <PlantHeadDashboard {...props} />} /> 
          <Route exact path="/ReportDashboard" render={() => <ReportDashboard {...props} />} />
          <Route exact path="/Reports" render={() => <Reports {...props} />} /> 
        </Switch>
      </div>
    </div>
  );
};


const Drr: React.FC<ISonaDeviationProps> = (props) => {
  return (
    <Router>
      <Layout {...props} />
    </Router>
  );
};

export default Drr;
