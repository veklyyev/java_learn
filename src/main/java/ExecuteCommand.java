import java.io.IOException;
public class ExecuteCommand {
    /**
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException {
        try {
            String FileName = "c:\\temp\\Digital_R2_MSA_Import_MD_4261.xlsm";
            String PSCommand = "c:\\windows\\SysWOW64\\cscript.exe c:\\Users\\Ihor_Petruniv\\IdeaProjects\\LearnJava\\excel_modify.vbs " + FileName;
            Process Result = Runtime.getRuntime().exec(PSCommand);
            Result.waitFor();
            System.out.print(Result.exitValue());
        }
        catch (Exception err){
            err.printStackTrace();
        }
    }

}
