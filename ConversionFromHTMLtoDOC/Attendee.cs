using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConversionFromHTMLtoDOC
{
    public class Attendee
    {
        public int Id { get; set; }
        public int MeetingId { get; set; }
        public string Name { get; set; }
        public string Company { get; set; }
    }
}
