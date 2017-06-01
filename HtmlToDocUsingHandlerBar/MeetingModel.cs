using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HtmlToDocUsingHandlerBar
{
    public class ActionItem
    {
        public int Id { get; set; }
        public int MeetingActionId { get; set; }
        public string Text { get; set; }
    }

    public class MeetingAction
    {
        public int Id { get; set; }
        public int ActionTypeId { get; set; }
        public int MeetingId { get; set; }
        public string Discussion { get; set; }
        public string Conclusion { get; set; }
        public List<ActionItem> ActionItems { get; set; }
        public string ActionName { get; set; }
    }

    public class MeetingModel
    {
        public int MeetingId { get; set; }
        public int AssociationId { get; set; }
        public string Type { get; set; }
        public string GroupDecision { get; set; }
        public string LoggedBy { get; set; }
        public string ModifiedBy { get; set; }
        public string CreatedDate { get; set; }
        public string ModifiedDate { get; set; }
        public string PerformancePoints { get; set; }
        public object Ytd { get; set; }
        public string Daa { get; set; }
        public string Zone { get; set; }
        public string Nation { get; set; }
        public bool IsFinalized { get; set; }
        public string District { get; set; }
        public bool IsDeleted { get; set; }
        public List<MeetingAction> MeetingActions { get; set; }
        public string AssociationName { get; set; }
        public string Make { get; set; }
        public string ThemeName { get; set; }
        public string ImageUrl { get; set; }
        public object Attendees { get; set; }
    }
}
