When we run a test in QTP, QTP generates a report based on inbuilt native reporting for steps  defined in QTP. To add user defined results, we have reporter object in QTP.
Below are the properties  and methods  of reporter object which will be described in details below.

Reporter.Filter Property 
Retrieves or sets the current mode for displaying events in the Test Results. You can use this property to completely disable or enable reporting of steps following the statement, or you can indicate that you only want subsequent failed or failed and warning steps to be included in the report.

Reporter.filter = FilterValue.

Filter value can be
·         0 or rfEnableAll – Default , all reporting events are displayed.
·         1 or rfEnableErrorandWarnings – only error and warning events are displayed.
·         2 or rfEnableErrorsOnly – Only errors are displayed.
·         3 or rfDisableAll – No event are displayed in results.

Reporter.ReportEvent Method
Reports an event to the results
Syntax: Reporter.reportEvent  eventstatus, reportStepName, Details,[Image]

EventStatus can have value 0, 1,2,3 or micPass,micFail, micDone, and micwarning.
reportStepName – Name of the steps
Details – Details of execution.
Image – Path of imagefile to be attached for event.

Reporter.ReportPath property
 This retrieves the folder Path in which current test results are saved.

Reporter.RunStatus property
Retrieves the run status at the current point of the run session. For tests, it returns the status of current test during the test run.
