using System;

namespace NYDOE.PMO.PopulateTasks.Entities
{
    public class UserTask
    {
        public int UID { get; set; }
        public string TaskID { get; set; }
        public string TaskName { get; set; }
        public string TaskDescription { get; set; }
        public string TaskUrl { get; set; }
        public string TaskStatus { get; set; }
        public bool IsComplete { get; set; }
        public decimal PercentComplete { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime DueDate { get; set; }
        public string Notes { get; set; }
    }
}
