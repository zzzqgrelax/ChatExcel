using ChatExcel.Addin.RTD;
using ChatExcel.Shared.Consts;
using ChatExcel.Shared.Models;
using NamedPipeClient;
using NamePipeLib;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChatExcel.Addin.Utilties
{
    static class NamedPipeServerHelper
    {
        internal static bool ServerStarted = false;
        public static void InitServer()
        {
            if (ServerStarted)
                return;

            ServerStarted = true;

            var ReceiveDataServer = new NamedPipeServer<string>(NamedPipeConst.ReceiveDataPipe);
            ReceiveDataServer.ClientMessage += delegate (NamedPipeConnection<string, string> conn, string message)
            {
                if (!string.IsNullOrEmpty(message))
                {
                    var data = JsonConvert.DeserializeObject<RTDDataDto>(message);
                    var dto = new RTDDataDto
                    {
                        Ticks = data.Ticks,
                        DicDatas = data.DicDatas,
                        ErrorMsg = data.ErrorMsg
                    };
                    RTDData.Datas.Add(dto);

                    if (!string.IsNullOrEmpty(data.ErrorMsg))
                        StatusBarMsgHelper.Info(data.ErrorMsg);
                }
            };
            ReceiveDataServer.Start();
        }
    }
}
