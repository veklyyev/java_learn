import java.io.IOException;

public class ExecuteCommand {
    /**
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException {
        try {
            String CurrentDir = System.getProperty("user.dir");
            String FileName =  CurrentDir + "\\Digital_R2_MSA_Import_MD_4261.xlsm";
            String PSCommand = "c:\\windows\\SysWOW64\\cscript.exe " + "C:\\jenkins\\workspace\\Cutom_Test\\src\\test\\resources\\openVBScript.vbs";
            Process Result = Runtime.getRuntime().exec(PSCommand);
            Result.waitFor();
            System.out.print(Result.exitValue());
        }
        catch (Exception err){
            err.printStackTrace();
        }
    }

}
