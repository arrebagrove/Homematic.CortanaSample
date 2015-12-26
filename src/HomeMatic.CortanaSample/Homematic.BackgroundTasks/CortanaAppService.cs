using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel.AppService;
using Windows.ApplicationModel.Background;
using Windows.ApplicationModel.VoiceCommands;
using Windows.Web.Http;

namespace Homematic.BackgroundTasks
{
    public sealed class CortanaAppService : IBackgroundTask
    {
        private BackgroundTaskDeferral serviceDeferral;
        VoiceCommandServiceConnection voiceServiceConnection;

        public async void Run(IBackgroundTaskInstance taskInstance)
        {
            //Take a service deferral so the service isn't terminated
            this.serviceDeferral = taskInstance.GetDeferral();

            taskInstance.Canceled += OnTaskCanceled;

            var triggerDetails =
              taskInstance.TriggerDetails as AppServiceTriggerDetails;

            if (triggerDetails != null &&
              triggerDetails.Name == "CortanaAppService")
            {
                try
                {
                    voiceServiceConnection =
                      VoiceCommandServiceConnection.FromAppServiceTriggerDetails(
                        triggerDetails);

                    voiceServiceConnection.VoiceCommandCompleted +=
                      VoiceCommandCompleted;

                    VoiceCommand voiceCommand = await
                    voiceServiceConnection.GetVoiceCommandAsync();

                    switch (voiceCommand.CommandName)
                    {
                        case "toggleSwitch":
                        {
                            var pref =
                                voiceCommand.Properties["pref"][0];
                            var switchname =
                                voiceCommand.Properties["switch"][0];
                            var state =
                                voiceCommand.Properties["state"][0];
                            await ToggleSwitch(pref, switchname, state);
                            break;
                        }

                        // As a last resort launch the app in the foreground
                        default:
                            LaunchAppInForeground();
                            break;
                    }
                }
                finally
                {
                    if (this.serviceDeferral != null)
                    {
                        //Complete the service deferral
                        this.serviceDeferral.Complete();
                    }
                }
            }
        }

        private async Task ToggleSwitch(string pref, string switchname, string state)
        {
            HttpClient client = new HttpClient();
            var newvalue = state == "ein" || state == "on" ? "true" : "false";
            var uri =
                new Uri(String.Format("http://homematic-ccu2/config/xmlapi/statechange.cgi?ise_id=1643&new_value={0}",
                    newvalue));

            bool success = true;
            try
            {
                var result = await client.GetAsync(uri);
                if (result.IsSuccessStatusCode)
                {
                    // First, create the VoiceCommandUserMessage with the strings 
                    // that Cortana will show and speak.
                    var userMessage = new VoiceCommandUserMessage();
                    userMessage.DisplayMessage = String.Format("Ich habe {0} {1} {2} geschaltet", pref, switchname, state);
                    userMessage.SpokenMessage = String.Format("Ich habe {0} {1} {2} geschaltet", pref, switchname, state);

                    var response =
                        VoiceCommandResponse.CreateResponse(
                            userMessage);
                    await voiceServiceConnection.ReportSuccessAsync(response);
                }
                else
                {
                    success = false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                success = false;
            }

            if(!success)
            {
                var userMessage = new VoiceCommandUserMessage();
                userMessage.DisplayMessage = String.Format("Ich konnte {0} {1} nicht {2} schalten", pref, switchname,
                    state);
                userMessage.SpokenMessage = String.Format("Ich konnte {0} {1} nicht {2} schalten", pref, switchname,
                    state);
                var response =
                    VoiceCommandResponse.CreateResponse(
                        userMessage);
                await voiceServiceConnection.ReportFailureAsync(response);
            }
        }

        private void VoiceCommandCompleted(VoiceCommandServiceConnection sender, VoiceCommandCompletedEventArgs args)
        {
            if (this.serviceDeferral != null)
            {
                // Insert your code here
                //Complete the service deferral
                this.serviceDeferral.Complete();
            }
        }

        /// <summary>
        /// When the background task is cancelled, clean up/cancel any ongoing long-running operations.
        /// This cancellation notice may not be due to Cortana directly. The voice command connection will
        /// typically already be destroyed by this point and should not be expected to be active.
        /// </summary>
        /// <param name="sender">This background task instance</param>
        /// <param name="reason">Contains an enumeration with the reason for task cancellation</param>
        private void OnTaskCanceled(IBackgroundTaskInstance sender, BackgroundTaskCancellationReason reason)
        {
            System.Diagnostics.Debug.WriteLine("Task cancelled, clean up");
            if (this.serviceDeferral != null)
            {
                //Complete the service deferral
                this.serviceDeferral.Complete();
            }
        }

        /// <summary>
        /// Provide a simple response that launches the app. Expected to be used in the
        /// case where the voice command could not be recognized (eg, a VCD/code mismatch.)
        /// </summary>
        private async void LaunchAppInForeground()
        {
            var userMessage = new VoiceCommandUserMessage();
            userMessage.SpokenMessage = "Starting Homematic Cortana sample";

            var response = VoiceCommandResponse.CreateResponse(userMessage);

            response.AppLaunchArgument = "";

            await voiceServiceConnection.RequestAppLaunchAsync(response);
        }

    }
}
