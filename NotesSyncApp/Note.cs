using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace NotesSyncApp
{
    class Note  
    {
        public int Id { get; set; }

        public string Title { get; set; }

        public string Colour { get; set; }

        [JsonIgnore]
        public string Rtf { get; set; }

        [JsonIgnore]
        public string DesiredLocation { get; set; }

        [JsonIgnore]
        public int Length { get; set; }

        public DateTime LastModified { get; set; }

        public string LastModifiedBy { get; set; }


        public Note()
        {
        }

        public Note(int id)
        {
            Id = id;
        }

        public Note(int id, string desiredLocation)
        {
            Id = id;
            DesiredLocation = desiredLocation;
        }

    }
}
