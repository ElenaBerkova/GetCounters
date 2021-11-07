using System;
using System.Collections.Generic;
using System.Linq;
using System.Reactive.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using ReactiveUI;

namespace GetCountersLib
{
    public class Starter
    {
        private readonly ExcelWorker _excel;
        private readonly CountersWorker _countWorker;
        private readonly List<CountersModel> _models;
        private readonly DateTime _startTime;
        private readonly DateTime _stopTime;
        private readonly long _timePeriod;
        private readonly long _timeCapture;
        private readonly long _timeCapturePeriod;

        public Starter(long timePeriodMinute = 40 , long timeCaptureSecond = 180, long timeCapturePeriodSecond = 5)
        {
            _timePeriod = timePeriodMinute * 60;
            _timeCapture = timeCaptureSecond;
            _timeCapturePeriod = timeCapturePeriodSecond;

            _excel = new ExcelWorker();
            _countWorker = new CountersWorker();
            _models = new List<CountersModel>();
            _startTime = new DateTime(2021, 12, 12, 8, 0, 0);
            _stopTime = new DateTime(2021, 12, 12, 17, 0, 0);
        }

        public async Task AddCountersAsync()
        {
            var cts = new CancellationTokenSource(TimeSpan.FromSeconds(_timeCapture));
            await Task.Run(async () =>
            {
                Console.WriteLine("First");
                _models.Clear();
                while (true)
                {
                    if(cts.IsCancellationRequested)
                    {
                        PrintResult();
                        _excel.AddToExcelFile(_models);
                        return;
                    }

                    _models.Add(new CountersModel
                    {
                        DateTime = DateTime.Now,
                        CpuCounter = _countWorker.GetCpuCounter(),
                        RamCounter = _countWorker.GetRamCounter(),
                        DiskCounter = _countWorker.GetDiskCounter()
                    });
                    await Task.Delay(TimeSpan.FromSeconds(_timeCapturePeriod)).ConfigureAwait(false);
                }
            }, cts.Token).ConfigureAwait(false);
        }
        public IDisposable CreateObservable() =>
            Observable.Interval(TimeSpan.FromSeconds(_timePeriod))
                .Where(_ => _startTime.Hour < DateTime.Now.Hour && _stopTime.Hour > DateTime.Now.Hour)
                .Subscribe(async next =>
                {
                    Console.WriteLine("ob start");
                    await AddCountersAsync().ConfigureAwait(false);
                });

        public void CloseExcelFile()
        {
            try
            {
                _excel.CloseExcelFile();
                Console.WriteLine("Closed");
            }
            catch (Exception ex)
            {
                // ignored
            }
        }

        public void PrintResult()
        {
            foreach(var model in _models) Console.WriteLine(model);
        }
    }
}
