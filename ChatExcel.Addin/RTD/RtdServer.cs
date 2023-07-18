using ChatExcel.Shared;
using ExcelDna.Integration.Rtd;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace ChatExcel.Addin.RTD
{
    [ProgId(ServerProgId)]
    public class RtdServer : ExcelRtdServer
    {
        public const string ServerProgId = "ChatExcel.Server.1";

        public static DataService _dataService;

        protected override bool ServerStart()
        {
            if (_dataService == null)
                _dataService = new DataService();
            return true;
        }

        protected override void ServerTerminate()
        {
            if (_dataService == null)
                return;

            if (_dataService.topics != null && _dataService.topics.Count > 0)
                return;

            _dataService.Terminate();
        }

        protected override Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            var sheetIdString = topicInfo.LastOrDefault();
            topicInfo = topicInfo.Take(topicInfo.Count - 1).ToList();
            var paras = new List<string>();
            paras.AddRange(topicInfo.ToList());

            var item = new KeyValuePair<string, RealRTDData>(
                       key: topicId.ToString(),
                       value: new RealRTDData()
                       {
                           Params = paras,
                           Value = Common.TopicDefaultLoading,
                           SheetId = new IntPtr(Convert.ToInt64(sheetIdString))
                       });

            _dataService.topics.Add(item.Key, item.Value);

            return base.CreateTopic(topicId, topicInfo);
        }


        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            try
            {
                //如果要使用缓存值 设置为false 
                //另外需要使用稳定注册 参考UDF里
                newValues = true;

                if (topicInfo.Count > 0)
                {
                    var paras = new List<string>();
                    paras.AddRange(topicInfo.ToList());

                    RealRTDData existing = null;
                    var existed = _dataService.topics.TryGetValue(topic.TopicId.ToString(), out existing);
                    if (existed)
                        existing.Topic = topic;

                    if (paras != null && paras.Count > 0
                       && (paras.Any(t => t.Equals(Common.TopicDefaultLoading))
                       || paras.Any(g => g.Equals(Common.TopicDefaultValue))))
                    {
                        existing.Value = Common.TopicDefaultValue;
                        topic.UpdateValue(Common.TopicDefaultValue);
                        return Common.TopicDefaultValue;
                    }

                    _dataService.ConnectTopic();

                    topic.UpdateValue(Common.TopicDefaultLoading);
                    return Common.TopicDefaultLoading;
                }
                else
                {
                    topic.UpdateValue(Common.TopicDefaultValue);
                    return Common.TopicDefaultValue;
                }
            }
            catch (Exception)
            {
                topic.UpdateValue(Common.TopicDefaultValue);
                return Common.TopicDefaultValue;
            }
        }

        protected override void DisconnectData(Topic topic)
        {
            _dataService.DisconnectTopic(topic);
        }
    }
}
