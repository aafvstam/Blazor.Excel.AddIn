/*
 * Copyright (c) Maarten van Stam. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 *
 */

/// <reference path="./blazor.d.ts" />

console.log("Loading command.js");

// TODO: Make a generic function to load the taskpanes.
// -----------------------------------------------------
//async function startAction(event: Office.AddinCommands.Event) {
//  try {
//    console.log("Running " + event.source.id);

//    switch (event.source.id) {
//      case "Contoso.Taskpane.Weather.Button":
//        console.log("Navigating to Weather page via Blazor router");

//        // First ensure DotNet is loaded
//        const dotnetLoaded = await preloadDotNet();

//        if (dotnetLoaded) {
//          // Use Blazor's NavigationManager to navigate programmatically
//          await DotNet.invokeMethodAsync(
//            "Blazor.Excel.AddIn.Client",
//            "NavigateToPage",
//            "/weather"
//          );
//          console.log("Navigation request to Weather page sent");
//        } else {
//          console.error("Could not load .NET runtime to navigate");
//        }
//        break;
//      case "Contoso.Taskpane.Counter.Button":
//        console.log("Navigating to Counter page via Blazor router");

//        // First ensure DotNet is loaded
//        const dotnetLoadedForCounter = await preloadDotNet();

//        if (dotnetLoadedForCounter) {
//          // Use Blazor's NavigationManager to navigate programmatically
//          await DotNet.invokeMethodAsync(
//            "Blazor.Excel.AddIn.Client",
//            "NavigateToPage",
//            "/counter"
//          );
//          console.log("Navigation request to Counter page sent");
//        } else {
//          console.error("Could not load .NET runtime to navigate");
//        }
//        break;
//      case "Contoso.Taskpane.Home.Button":
//        console.log("Navigating to Home page via Blazor router");

//        // First ensure DotNet is loaded
//        const dotnetLoadedForHome = await preloadDotNet();

//        if (dotnetLoadedForHome) {
//          // Use Blazor's NavigationManager to navigate programmatically
//          await DotNet.invokeMethodAsync(
//            "Blazor.Excel.AddIn.Client",
//            "NavigateToPage",
//            "/home"
//          );
//          console.log("Navigation request to Home page sent");
//        } else {
//          console.error("Could not load .NET runtime to navigate");
//        }
//        break;
//      default:
//        console.log("Unknown command: " + event.source.id);
//    }
//  }
//  catch (e: any) {
//    console.error("Error in startAction: " + e.message);
//  }
//  finally {
//    event.completed();
//    console.log("Finish " + event.source.id);
//  }
//}

/**
 * Writes the text from the Home Blazor Page to the Worksheet when highlightSelectionHome runs.
 * @param event {Office.AddinCommands.Event}
 */
async function highlightSelectionHome(event: Office.AddinCommands.Event) {

  // Implement your custom code here. The following code is a simple Excel example.  
  try {
    console.log("Running highlightSelectionHome");

    console.log("Before callStaticLocalComponentMethodinit");
    await callStaticLocalComponentMethodinit("SayHelloHome");
    console.log("After callStaticLocalComponentMethodinit");

    // Used to verify the previous function call, if that fails, this will not run.
    // It will be skipped on error and jump into the catch block.
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "LightBlue";
      await context.sync();
    });

  } catch (error: any) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }
  finally {
    console.log("Finish highlightSelectionHome");

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
  }
}

/**
 * Writes the text from the Counter Blazor Page to the Worksheet when highlightSelectionCounter runs.
 * @param event {Office.AddinCommands.Event}
 */
async function highlightSelectionCounter(event: Office.AddinCommands.Event) {

  // Implement your custom code here. The following code is a simple Excel example.  
  try {
    console.log("Running highlightSelectionCounter");

    console.log("Before callStaticLocalComponentMethodinit");
    await callStaticLocalComponentMethodinit("SayHelloCounter");
    console.log("After callStaticLocalComponentMethodinit");

    // Used to verify the previous function call, if that fails, this will not run.
    // It will be skipped on error and jump into the catch block.
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "LightBlue";
      await context.sync();
    });

  } catch (error: any) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }
  finally {
    console.log("Finish highlightSelectionHome");

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
  }
}

/**
 * Writes the event source id to the document when ExecuteFunction runs.
 * @param event {Office.AddinCommands.Event}
 */
async function writeValue(event: Office.AddinCommands.Event) {

  console.log("In writeValue");

  try {
    let message = "ExecuteFunction works. Button ID=" + event.source.id;

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.values = [[message]];
      range.getEntireColumn().format.autofitColumns();
      await context.sync();
    });

    console.log("writeValue Succeeded");

  } catch (error: any) {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const cellRange = range.getCell(0, 0);
      cellRange.values = [[error.message]];
      await context.sync();
    });
    console.log();
    console.log("Error call : " + error.message);
  }
  finally {
    console.log("Finish writeValue");

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
  }
}

/**
 * Calls the JSInvokable function CreateBubbles to create a bubble chart, after adding the data in the active worksheet.
 * @param event {Office.AddinCommands.Event}
 */
async function createBubbles(event: Office.AddinCommands.Event) {

  console.log("Running createBubbles");

  // Implement your custom code here. The following code is a simple Excel example.
  try {

    // Call JSInvokable Function here ...
    await DotNet.invokeMethodAsync(
      "Blazor.Excel.AddIn.Client",
      "CreateBubbles");

    console.log("Finished createBubbles")

  } catch (error: any) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }
  finally {
    console.log("Finish createBubbles");
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
  }
}

/**
 * Writes the text from the Home Blazor Page to the Excel Cell
 * @param {any} event
 */
async function callBlazorOnHome(event: Office.AddinCommands.Event) {
  // Implement your custom code here. The following code is a simple Excel example.
  try {
    console.log("Running callBlazorOnHome");

    console.log("Before callStaticLocalComponentMethodinit");
    await callStaticLocalComponentMethodinit("SayHelloHome");
    console.log("After callStaticLocalComponentMethodinit");
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  } finally {
    console.log("Finish callBlazorOnHome");

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
  }
}

/**
 * Writes the text from the Counter Blazor Page to the Excel Cell
 * @param {any} event
 */
async function callBlazorOnCounter(event: Office.AddinCommands.Event) {
  // Implement your custom code here. The following code is a simple Excel example.
  try {
    console.log("Running callBlazorOnHome");

    console.log("Before callStaticLocalComponentMethodinit");
    await callStaticLocalComponentMethodinit("SayHelloCounter");
    console.log("After callStaticLocalComponentMethodinit");
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  } finally {
    console.log("Finish callBlazorOnCounter");

    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
  }

}

/**
 * Checks if the .NET runtime is loaded and invokes a .NET method to retrieve a string.
 * The string is then inserted into a Excel Cell.
 *
 * @param {string} methodname - The name of the .NET method to invoke.
 */
async function callStaticLocalComponentMethodinit(methodname: string) {
  console.log("In callStaticLocalComponentMethodinit");

  try {
    let name = "Initializing";

    try {
      var dotnetloaded = await preloadDotNet();

      if (dotnetloaded === true) {

        // Call JSInvokable Function here ...
        name = await DotNet.invokeMethodAsync(
          "Blazor.Excel.AddIn.Client",
          methodname,
          "Blazor Fan"
        );

      } else {
        name = "Init DotNet Failed";
      }
    } catch (error: any) {
      name = error.message;
      console.error("Error during DotNet invocation: " + name);
    }

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.values = [[name]];
      range.getEntireColumn().format.autofitColumns();
      await context.sync();
    });

    // Used to verify the previous function call, if that fails, this will not run.
    // It will be skipped on error and jump into the catch block.
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "yellow";
      await context.sync();
    });
  }
  catch (error: any) {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const cellRange = range.getCell(0, 0);
      cellRange.values = [[error.message]];
      await context.sync();
    });

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "red";
      await context.sync();
    });

    console.log();
    console.log("Error call : " + error.message);
  }
  finally {
    console.log("Finish callStaticLocalComponentMethodinit");
  }
}

/**
 * Local function to preload the .NET runtime and ensure it is ready for use.
 *
 * This function attempts to invoke a dummy method in the .NET runtime to check if it is loaded.
 * It retries up to 5 times with a 1-second delay between attempts if the runtime is not loaded.
 *
 * This won't be necessary if the task pane is automatically opened when the add-in is loaded.
 * Also feel it should be possible to preload in the module.ts file for a guaranteed load.
 *
 * @returns result - Returns true if the .NET runtime is successfully loaded, otherwise false.
 */
async function preloadDotNet() {
  console.log("In preloadDotNet");
  try {
    console.log("Before DotNet.invokeMethodAsync");
    var result = "";

    let attempts = 0;
    while (result === "" && attempts < 5) {
      try {
        if (attempts > 0) {
          await new Promise((resolve) => setTimeout(resolve, 1000));
        }
        result = await DotNet.invokeMethodAsync(
          "Blazor.Excel.AddIn.Client",
          "PreloaderDummy"
        );
      } catch (error: any) {
        console.error("Error during DotNet invocation: " + error.message);
      }
      attempts++;
    }

    return result === "Loaded" ? true : false;

    console.log("After DotNet.invokeMethodAsync");
  } catch (error: any) {
    console.log("Error call : " + error.message);
  } finally {
    console.log("Finish preloadDotNet");
  }

  return false;
}

// Associate the functions with their named counterparts in the manifest XML.
// Office.actions.associate("startAction", startAction);
Office.actions.associate("highlightSelectionHome", highlightSelectionHome);
Office.actions.associate("highlightSelectionCounter", highlightSelectionCounter);
Office.actions.associate("writeValue", writeValue);
Office.actions.associate("createBubbles", createBubbles);
Office.actions.associate("callBlazorOnHome", callBlazorOnHome);
Office.actions.associate("callBlazorOnCounter", callBlazorOnCounter);
