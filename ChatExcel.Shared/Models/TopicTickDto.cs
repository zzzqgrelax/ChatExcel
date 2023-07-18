using System.Collections.Generic;

namespace ChatExcel.Shared.Models
{
    public class TopicTickDto
    {
        public long Ticks { get; set; }

        public List<TopicParam> Topics { get; set; }
    }

    public class TopicParam
    {
        public string TopicId { get; set; }
        public List<string> Params { get; set; }
    }

    public class TopicDataDto
    {
        public string TopicId { get; set; }

        public string Value { get; set; }
    }

    public class RTDDataDto
    {
        public Dictionary<string, TopicDataDto> DicDatas { get; set; } = new Dictionary<string, TopicDataDto>();

        public long Ticks { get; set; }

        public string ErrorMsg { get; set; }
    }
}
