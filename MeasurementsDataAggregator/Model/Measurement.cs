using System;

namespace MeasurementsDataAggregator.Model
{
    class Measurement
    {
        public DateTime Date { get; set; }
        public float Weight { get; set; }
        public float MuscleMass { get; set; }
        public float BMI { get; set; }
        public float VisceralFatIndex { get; set; }
        public float FatMassPercentage { get; set; }
    }
}
