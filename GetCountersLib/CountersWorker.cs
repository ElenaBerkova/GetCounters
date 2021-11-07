using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetCountersLib
{
    public class CountersWorker
    {
        private readonly PerformanceCounter _cpuCounter;
        private readonly PerformanceCounter _ramCounter;
        private readonly PerformanceCounter _diskCounter;

        public CountersWorker()
        {
            _cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
            _ramCounter = new PerformanceCounter("Memory", "% Committed Bytes In Use", null);
            _diskCounter = new PerformanceCounter("PhysicalDisk", "Avg. Disk Queue Length", "_Total");
        }

        public float GetCpuCounter() => _cpuCounter.NextValue();
        public float GetRamCounter() => _ramCounter.NextValue();
        public float GetDiskCounter() => _diskCounter.NextValue();
    }
}
