using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetCountersLib
{
    public class CountersModel
    {
        public DateTime DateTime { get; set; }
        public float CpuCounter { get; set; }
        public float RamCounter { get; set; }
        public float DiskCounter { get; set; }

        public override string ToString() => new StringBuilder()
            .Append($"{DateTime}\t")
            .Append($"{CpuCounter}\t")
            .Append($"{RamCounter}\t")
            .Append($"{DiskCounter}\t")
            .ToString();
    }
}
