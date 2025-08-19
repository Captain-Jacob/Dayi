namespace ProgramKontrol;

static class Program    // buraya dokunmadın ben ,bilginize
{
  
    [STAThread]
    static void Main()
    {
 
        ApplicationConfiguration.Initialize();
        Application.Run(new Form1());
    }    
}