function PerformStart() {
   // Do any initialization.  This happens when the alert is initialized, not when the firing starts.
}

function PerformStop() {

}

function TestAlert(alert, collector, host) {

    ExecuteProcess(2000);
}

function PerformAction(state, didChange) {
   ExecuteProcess();
}

function ExecuteProcess(waitTime) {
   if (!IsLaunchExeAllowed()) {
      return;
   }
   var newProcess = new System.Diagnostics.Process();
   newProcess.StartInfo.UseShellExecute = true;
   newProcess.StartInfo.Arguments = ParseTemplate(Action.Settings.Parameters);
   newProcess.StartInfo.FileName = Action.Settings.ProcessName;
   newProcess.Start();
   if (waitTime) {
      newProcess.WaitForExit(waitTime);
   }
}