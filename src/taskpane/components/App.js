import * as React from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList from "./HeroList";
import Progress from "./Progress";

/* global console, Excel, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  // click = async () => {
  //   try {
  //     await Excel.run(async (context) => {
  //       // Create a new worksheet named "RAW"
  //       const rawSheet = context.workbook.worksheets.add("RAW");
  //       rawSheet.visibility = "Hidden"

  //       // Insert your Excel code here
  //       const response = [
  //         {
  //           name: "Atharva"
  //         },
  //         {
  //           name: "Prasanna"
  //         },
  //         {
  //           name: "Someone"
  //         }
  //       ];

  //       // Get the range where the values should be inserted in RAW sheet
  //       const startCell = rawSheet.getRange("A1");
  //       const endCell = startCell.getOffsetRange(response.length - 1, 0);
  //       const range = startCell.getBoundingRect(endCell);

  //       // Insert the values from the response array into the column in RAW sheet
  //       for (let i = 0; i < response.length; i++) {
  //         range.getCell(i, 0).values = [[response[i].name]];
  //       }

  //       // Update the fill color of the range in RAW sheet
  //       range.format.fill.color = "yellow";

  //       // Get the range containing the names in RAW sheet
  //       const rawNamesRange = rawSheet.getRange("A1:A" + response.length);

  //       // Load the names range to sync with the workbook
  //       rawNamesRange.load("values");

  //       // Synchronize the changes with the workbook
  //       await context.sync();

  //       // Get the first worksheet (Sheet1)
  //       const sheet1 = context.workbook.worksheets.getActiveWorksheet();

  //       // Get the range where the dropdown list should be added in Sheet1
  //       const dropdownRange = sheet1.getRange("A1");

  //       // Set data validation to create a dropdown list with values from the RAW sheet
  //       dropdownRange.dataValidation.rule = {
  //         list: {
  //           inCellDropDown: true, // Show the dropdown always
  //           source: rawNamesRange
  //         },
  //         formula1: "=RAW!$A$1:$A$" + response.length
  //       };

  //       await context.sync();
  //       console.log(`The range address was ${dropdownRange.address}.`);
  //     });
  //   } catch (error) {
  //     console.error(error);
  //   }
  // };


  // click = async () => {
  //   try {
  //     await Excel.run(async (context) => {
  //       // Create a new worksheet named "RAW"
  //       const rawSheet = context.workbook.worksheets.add("RAW");
  
  //       // Insert your Excel code here
  //       const response = [
  //         {
  //           name: "Atharva"
  //         },
  //         {
  //           name: "Prasanna"
  //         },
  //         {
  //           name: "Someone"
  //         }
  //       ];
  
  //       // Get the range where the values should be inserted in RAW sheet
  //       const startCell = rawSheet.getRange("A1");
  //       const endCell = startCell.getOffsetRange(response.length - 1, 0);
  //       const range = startCell.getBoundingRect(endCell);
  
  //       // Insert the values from the response array into the column in RAW sheet
  //       for (let i = 0; i < response.length; i++) {
  //         range.getCell(i, 0).values = [[response[i].name]];
  //       }
  
  //       // Update the fill color of the range in RAW sheet
  //       range.format.fill.color = "yellow";
  
  //       // Get the range containing the names in RAW sheet
  //       const rawNamesRange = rawSheet.getRange("A1:A" + response.length);
  
  //       // Load the names range to sync with the workbook
  //       rawNamesRange.load("values");
  
  //       // Synchronize the changes with the workbook
  //       await context.sync();
  
  //       // Get the first worksheet (Sheet1)
  //       const sheet1 = context.workbook.worksheets.getActiveWorksheet();
  
  //       // Get the range where the dropdown lists should be added in Sheet1 for the first line (A1) and second line (A2)
  //       const dropdownRange1 = sheet1.getRange("A1");
  //       const dropdownRange2 = sheet1.getRange("A2");
  
  //       // Set data validation to create a dropdown list with values from the RAW sheet
  //       dropdownRange1.dataValidation.rule = {
  //         list: {
  //           inCellDropDown: true, // Show the dropdown always
  //           source: rawNamesRange
  //         },
  //         formula1: "=RAW!$A$1:$A$" + response.length
  //       };
  
  //       dropdownRange2.dataValidation.rule = {
  //         list: {
  //           inCellDropDown: true, // Show the dropdown always
  //           source: rawNamesRange
  //         },
  //         formula1: "=RAW!$A$1:$A$" + response.length
  //       };
  
  //       // Get the range in column C where the selected values should be displayed
  //       const selectedValuesRange = sheet1.getRange("C1:C2");
  
  //       // Get the selected values from the dropdowns in column A and set them in column C
  //       selectedValuesRange.formulas = [
  //         ["=A1"],
  //         ["=A2"]
  //       ];
  
  //       await context.sync();
  //       console.log(`Dropdowns created and selected values displayed in column C.`);
  //     });
  //   } catch (error) {
  //     console.error(error);
  //   }
  // };
  
  
  click = async () => {
    try {
      await Excel.run(async (context) => {
        // Create a new worksheet named "RAW"
        const rawSheet = context.workbook.worksheets.add("RAW");
  
        // Insert your Excel code here
        const response = [
          {
            name: "Atharva"
          },
          {
            name: "Prasanna"
          },
          {
            name: "Someone"
          },
          {
            name: "Ashish"
          }
        ];
  
        // Get the range where the values should be inserted in RAW sheet
        const startCell = rawSheet.getRange("A1");
        const endCell = startCell.getOffsetRange(response.length - 1, 0);
        const range = startCell.getBoundingRect(endCell);
  
        // Insert the values from the response array into the column in RAW sheet
        for (let i = 0; i < response.length; i++) {
          range.getCell(i, 0).values = [[response[i].name]];
        }
  
        // Update the fill color of the range in RAW sheet
        range.format.fill.color = "yellow";
  
        // Get the range containing the names in RAW sheet
        const rawNamesRange = rawSheet.getRange("A1:A" + response.length);
  
        // Load the names range to sync with the workbook
        rawNamesRange.load("values");
  
        // Synchronize the changes with the workbook
        await context.sync();
  
        // Get the first worksheet (Sheet1)
        const sheet1 = context.workbook.worksheets.getActiveWorksheet();
  
        // Number of dropdowns to create (assuming response.length is the desired number)
        const numberOfDropdowns = response.length;
  
        // Create dropdowns and set data validation dynamically
        for (let i = 0; i < numberOfDropdowns; i++) {
          const dropdownRange = sheet1.getRange(`A${i + 1}`);
          dropdownRange.dataValidation.rule = {
            list: {
              inCellDropDown: true, // Show the dropdown always
              source: rawNamesRange
            }
          };
        }
  
        // Synchronize the changes to show dropdowns
        await context.sync();
  
        // Get the range in column C where the selected values should be displayed
        const selectedValuesRange = sheet1.getRange(`C1:C${numberOfDropdowns}`);
  
        // Create formulas to link the selected values from the dropdowns to column C
        const formulas = [];
        for (let i = 0; i < numberOfDropdowns; i++) {
          formulas.push([`=A${i + 1}`]);
        }
  
        // Set the formulas to link the selected values to column C
        selectedValuesRange.formulas = formulas;
  
        await context.sync();
        console.log(`Dropdowns created and selected values displayed in column C.`);
      });
    } catch (error) {
      console.error(error);
    }
  };
  



  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Run
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
