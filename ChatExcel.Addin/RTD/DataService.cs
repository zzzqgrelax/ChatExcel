using ChatExcel.Addin.Ribbon;
using ChatExcel.Addin.Utilties;
using ChatExcel.Shared;
using ChatExcel.Shared.Consts;
using ChatExcel.Shared.Models;
using ExcelDna.Integration.Rtd;
using NamePipeLib;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ChatExcel.Addin.RTD
{
    public class DataService
    {
        private Thread _updateThread;
        internal Dictionary<string, RealRTDData> topics = new Dictionary<string, RealRTDData>();
        internal bool _isBusy;
        internal bool _fetching;

        public DataService()
        {
            _updateThread = new Thread(RunUpdates);
            _updateThread.Start();
            topics = new Dictionary<string, RealRTDData>();
        }

        public void ConnectTopic()
        {
            lock (topics)
            {
                _fetching = true;
                if (!_updateThread.IsAlive)
                {
                    _updateThread = new Thread(RunUpdates);
                    _updateThread.Start();
                }
            }
        }

        public void DisconnectTopic(ExcelRtdServer.Topic topic)
        {
            lock (topics)
            {
                try
                {
                    topics.Remove(topic.TopicId.ToString());
                }
                catch (Exception)
                { }
            }
        }

        public void Terminate()
        {
            try
            {
                _updateThread.Abort();
            }
            catch (Exception)
            { }
        }


        void RunUpdates()
        {
            try
            {
                while (true)
                {
                    UpdateSomeTopics();
                    Thread.Sleep(1000);
                }
            }
            catch { }
        }

        bool IsTerminalExisted()
        {
            var processExisted = Process.GetProcesses()
                   .Any(pr => pr.ProcessName.ToLower().Equals(ProcessConst.ChatExcelProcess.ToLower()));
            return processExisted;
        }

        void UpdateSomeTopics()
        {
            try
            {
                if (_isBusy)
                    return;
                _isBusy = true;

                if (!_fetching)
                    return;

                lock (topics)
                {
                    if (!IsTerminalExisted())
                        return;
                    if (!RibbonController.LoggedIn)
                    {
                        StatusBarMsgHelper.Info(StatusBarMsgHelper.RTDNotLoginMsg);
                        return;
                    }

                    var fetchingTopics = topics.Where(t => t.Value.Value == null
                    || t.Value.Value.Equals(Common.TopicDefaultLoading))
                        .ToList();

                    if (fetchingTopics != null && fetchingTopics.Count > 0)
                        StatusBarMsgHelper.Info(StatusBarMsgHelper.RTDStartMsg);

                    var data = new TopicTickDto()
                    {
                        Ticks = DateTime.Now.Ticks,
                        Topics = new List<TopicParam>()
                    };

                    foreach (var itemTopic in fetchingTopics)
                    {
                        data.Topics.Add(new TopicParam() { Params = itemTopic.Value.Params,TopicId = itemTopic.Key});
                    }

                    if (data.Topics.Count > 0)
                    {
                        var client = new NamedPipeClient<string>(NamedPipeConst.QueryDataPipe);
                        client.PushMessage(JsonConvert.SerializeObject(data));
                        int i = 0;
                        while (!RTDData.Datas.Any(t => t.Ticks.Equals(data.Ticks)))
                        {
                            i++;
                            Task.Delay(200).Wait();
                            if (i > 3000)
                                break;
                        }
                    }

                    var existing = RTDData.Datas.FirstOrDefault(t => t.Ticks.Equals(data.Ticks));

                    foreach (KeyValuePair<string, RealRTDData> topic in fetchingTopics)
                    {
                        if (existing == null || existing.DicDatas == null)
                        {
                            if (topic.Value.Value == null || topic.Value.Value.ToString().Equals(Common.TopicDefaultLoading))
                                topic.Value.Value = Common.TopicDefaultValue;
                            continue;
                        }

                        if (topic.Value.Value != null
                            && topic.Value.Value.ToString() != Common.TopicDefaultLoading)
                            continue;

                        TopicDataDto info = null;
                        existing.DicDatas.TryGetValue(topic.Key.ToString(), out info);

                        if (info != null)
                        {
                            if (string.IsNullOrEmpty(info.Value) )
                            {
                                topic.Value.Value = Common.TopicDefaultValue;
                                continue;
                            }
                            topic.Value.Value = info.Value;
                        }
                        else
                        {
                            if (topic.Value.Value == null || topic.Value.Value.ToString().Equals(Common.TopicDefaultLoading))
                                topic.Value.Value = Common.TopicDefaultValue;
                        }
                    }

                    var server = topics.FirstOrDefault(t => t.Value.Topic != null).Value.Topic.Server;
                    server.UpdateValues(fetchingTopics.Where(t => t.Value.Topic != null).Select(t => t.Value.Topic).ToList(),
                    fetchingTopics.Select(t => t.Value.Value).ToList());

                    var updatedFetchingTopics = topics
                        .Where(t => t.Value.Value == null || t.Value.Value.Equals(Common.TopicDefaultLoading))
                       .ToList();

                    if (updatedFetchingTopics == null || updatedFetchingTopics.Count == 0)
                    {
                        var msg = existing?.ErrorMsg;
                        if (existing != null && !string.IsNullOrEmpty(existing.ErrorMsg))
                            StatusBarMsgHelper.Info(msg);
                        else
                            StatusBarMsgHelper.Info(StatusBarMsgHelper.RTDFinishMsg);

                        _fetching = false;
                    }

                    if (existing != null)
                    {
                        RTDData.Datas.Remove(existing);
                    }
                }
            }
            catch (Exception) { }
            finally
            {
                _isBusy = false;
            }
        }
    }
}
