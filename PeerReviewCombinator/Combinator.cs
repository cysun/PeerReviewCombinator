using Ascent.Helpers;
using Serilog;

namespace PeerReviewCombinator
{
    public class Settings
    {
        public string RosterFile { get; set; }
        public string InputFolder { get; set; }
        public string OutputFile { get; set; }
        public string[] ExpectedColumns { get; set; }
        public int OptionalColumnCount { get; set; }
    }

    public class Student
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Cin { get; set; }
    }

    public class Combinator
    {
        private readonly Settings _settings;

        private readonly Dictionary<string, Student> _roster;

        public Combinator(Settings settings)
        {
            _settings = settings;
            _roster = new Dictionary<string, Student>();
            loadRoster();
        }

        public void run()
        {
            var colNames = new List<string>() { "AssesorId", "CIN" };
            colNames.AddRange(_settings.ExpectedColumns);

            using var excelWriter = new ExcelWriter(_settings.OutputFile, colNames);

            var files = Directory.GetFiles(_settings.InputFolder);
            foreach (var file in files)
            {
                // The Canvas submission file has the following naming convention:
                //     <FullName>_<UserId>_<SubmissionId>_<OriginalFileName>
                var assessorId = Path.GetFileName(file).Split('_')[1];
                Log.Debug("Assessor Id: {assessorId}", assessorId);

                using var excelReader = new ExcelReader(file);
                if (!excelReader.HasColumns(_settings.ExpectedColumns))
                {
                    Log.Warning("{file} does not have all the expected columns", file);
                    continue;
                }

                while (excelReader.Next())
                {
                    if (excelReader.EmptyCellCount() > _settings.OptionalColumnCount) continue;

                    var studentName = excelReader.Get(0);
                    var newRow = new List<string>() { assessorId, _roster[studentName].Cin };
                    newRow.AddRange(excelReader.GetAll());
                    excelWriter.WriteRow(newRow);
                }
            }

            excelWriter.Finish();
        }

        private void loadRoster()
        {
            using var excelReader = new ExcelReader(_settings.RosterFile);
            while (excelReader.Next())
            {
                var student = new Student
                {
                    Id = int.Parse(excelReader.Get("ID")),
                    Name = excelReader.Get("Student"),
                    Cin = excelReader.Get("SIS User ID")
                };
                _roster[student.Name] = student;
            }
            Log.Information("Loaded {n} students", _roster.Count);
        }
    }
}
