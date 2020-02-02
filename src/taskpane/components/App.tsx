import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Progress from "./Progress";
/* global Button, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
    };
  }

  componentDidMount() {
   
  }

  sendClick = async () => {
      Office.context.ui.messageParent(JSON.stringify({ message: 'send' }))
  };

  cancelClick = async () => {
    Office.context.ui.messageParent(JSON.stringify({ message: 'cancel' }))
};

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.sendClick}
          >
            Send Mail
          </Button>

          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.cancelClick}
          >
            Cancel Mail
          </Button>

      </div>
    );
  }
}
