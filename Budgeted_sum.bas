<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabPlanner" label="Plaነer">
        <group id="grpP6" label="P6 Conversion Tools">
          <button id="btnWBSColoring"
                  label="WBS Coloring"
                  image="color-wheel"
                  size="large"
                  onAction="Planner_WBSColoring"
                  screentip="WBS Coloring"
                  supertip="Creates WBS Coloring Format up to 5 WBS level For Given Project "/>
          <button id="btnBudgetedSum"
                  label="Bugdeted Sum"
                  image="sum-sign"
                  size="large"
                  onAction="Planner_BudgetedSum"
                  screentip="Budgeted SUM Formula Giver"
                  supertip=" Gives Sum Formula For the WBS Part of the Schedule "/>
          <button id="btnARemover"
                  label="A remover"
                  image="A"
                  size="large"
                  onAction="Planner_ARemover"
                  screentip=" {A} Remover from the date "
                  supertip=" Remove { A } / Actual Start Date symbol from the Dates "/>
                  
          <button id="btnWeightage"
                  label="Weightage"
                  image="Weightage"
                  size="large"
                  onAction="Planner_Weightage"
                  screentip="Calculate the weightage Percentage"
                  supertip="Creates a Weightage formula means total budgeted value divided by the specific activities value ."/>
          <button id="btnPlannedPercentage"
                  label="Planned Percentage"
                  image="PlannedPercentage"
                  size="large"
                  onAction="Planner_PlannedPercentage"
                  screentip="Calculate Planned Percentage"
                  supertip="Creates a Planned Percentage column using Start Date, Finish Date, and a selected status date."/>
          <button id="btnPlannedValue"
                 label="Planned Value"
                 image="PlannedValue"
                 size="large"
                 onAction="Planner_PlannedValue"
                 screentip="Calculate Planned Value"
                 supertip="this formula will do By multipling planned percentage by the specific activties weightage gives us the planned value  "/>   
          <button id="btnActualPercentage"
                 label="Actual Percentage"
                 image="ActualPercentage"
                 size="large"
                 onAction="Planner_ActualPercentage"
                 screentip="Calculate Actual Percentage"
                 supertip="Creates a Actual Percentage column by dividing Actual quantity and Total quantity "/> 
         <button id="btnActualValue"
                label="Actual Value"
                image="ActualValue"
                size="large"
                onAction="Planner_ActualValue"
                screentip="Calculate Actual Value"
                supertip="Creates Actual Value column by multipling Actual percentage and Weightage"/>  
        </group>
 <group id="grpScheduleUpdate" label="Schedule Update">
  <button id="btnSyncByActivityID"
          label="Sync Qty/Budget"
          imageMso="RefreshAll"
          size="large"
          onAction="Planner_SyncByActivityID"
          screentip="Update by Activity ID"
          supertip="Updates Total Quantity and Budgeted Amt from another workbook using Activity ID as the reference."/>

<menu id="menuUpdateReview"
      label="Update Review"
      imageMso="ConditionalFormattingMenu">

  <button id="btnShowBlankActivities"
          label="Blank Activities - Yellow"
          imageMso="HighlightColorYellow"
          onAction="Planner_ShowBlankActivities"
          screentip="Highlight blank activities"
          supertip="Highlights blank activities in Yellow."/>

  <button id="btnShowMissingActivities"
          label="Missing Activities - Red"
          imageMso="ReviewDeleteComment"
          onAction="Planner_ShowMissingActivities"
          screentip="Highlight missing activities"
          supertip="Highlights missing activities in Red."/>

  <button id="btnShowDuplicateActivities"
          label="Duplicate Activities - Orange"
          imageMso="FileDocumentInspect"
          onAction="Planner_ShowDuplicateActivities"
          screentip="Highlight duplicate activities"
          supertip="Highlights duplicate activities in Orange."/>

  <button id="btnShowUpdatedRows"
          label="Updated Rows - Green"
          imageMso="AcceptInvitation"
          onAction="Planner_ShowUpdatedRows"
          screentip="Highlight updated rows"
          supertip="Highlights updated rows in Green."/>

  <button id="btnClearReviewColors"
          label="Clear Review Colors"
          imageMso="ClearFormats"
          onAction="Planner_ClearReviewColors"
          screentip="Clear review highlights"
          supertip="Removes all review highlight colors from the active sheet."/>

</menu>
</group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
