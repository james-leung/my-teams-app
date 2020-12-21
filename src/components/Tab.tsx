// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import React, { useEffect, useState } from "react";
import "./App.css";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * The 'PersonalTab' component renders the main tab content
 * of your app.
 */
export default  () => {
  const [context, setContext] = useState<microsoftTeams.Context>();
  let userName = "";
  if (context && "upn" in context) {
    userName = context["upn"] ?? "";
  }

  useEffect(() => {
    microsoftTeams.getContext((context: microsoftTeams.Context) => {
      setContext(context);
    });
  }, []);

  return (
    <div>
      <h3>Hello World!</h3>
      <h1>Congratulations {userName}!</h1>{" "}
      <h3>This is the tab you made :-)</h3>
    </div>
  );
};
