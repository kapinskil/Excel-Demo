namespace Excel_Demo
{
    class Program
    {
        static void Main(string[] args)
        {

            var excel = new Excel(@"C:\Users\kapin\OneDrive\Pulpit\test3.xlsx", 1);
            excel.CreateNewWorksheet();
            excel.SelectWorksheet(2);
            excel.WriteToCell(1,1,"This is sheet 2");
            excel.DeleteWorksheet(1);

            excel.SaveAs(@"C:\Users\kapin\OneDrive\Pulpit\test4.xlsx");
            excel.Close();

            
        }
    }
}
