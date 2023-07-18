using ChatExcel.Shared.Consts;
using ChatExcel.Shared.Models;
using NamedPipeClient;
using NamePipeLib;
using Newtonsoft.Json;
using Prism.DryIoc;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ChatExcel
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : PrismApplication
    {
        private string appName;
        public EventWaitHandle ProgramStarted { get; set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                appName = Path.GetFileNameWithoutExtension(Process.GetCurrentProcess().MainModule.FileName);
                ProgramStarted = new EventWaitHandle(false, EventResetMode.AutoReset, appName, out var createNew);
                if(createNew)
                {
                    base.OnStartup(e);
                    InitServer();
                }
            }
            catch { }
        }
        protected override Window CreateShell()
        {
            return Container.Resolve<MainWindow>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            
        }

        private void InitServer()
        {
            var queryDataServer = new NamedPipeServer<string>(NamedPipeConst.QueryDataPipe);
            queryDataServer.ClientMessage += async delegate (NamedPipeConnection<string, string> conn, string message)
            {
                if (!string.IsNullOrEmpty(message))
                {
                    await Task.Delay(1500);
                    var input = JsonConvert.DeserializeObject<TopicTickDto>(message);
                    var client = new NamedPipeClient<string>(NamedPipeConst.ReceiveDataPipe);
                    var data = new RTDDataDto() { Ticks = input.Ticks,DicDatas = new Dictionary<string, TopicDataDto>() };
                    foreach (var item in input.Topics)
                    {
                        //todo 
                        data.DicDatas.Add(key: item.TopicId, value: new TopicDataDto()
                        {
                            TopicId=item.TopicId,
                            Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ff")
                        });
                    }
                    client.PushMessage(JsonConvert.SerializeObject(data));
                }
            };
            queryDataServer.Start();
        }
     }
}
