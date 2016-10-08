import jxl.write.WriteException;
import org.opencv.core.Core;
import tw.edu.sju.ee.commons.nativeutils.NativeUtils;

import java.io.IOException;

/**
 * Created by Krzysiek on 2016-07-23.
 */
public class Application {
    static {
        try {
            System.loadLibrary(Core.NATIVE_LIBRARY_NAME); // used for tests. This library in classpath only
        } catch (UnsatisfiedLinkError e) {
            try {
                NativeUtils.loadLibraryFromJar("opencv_java310"); // during runtime. .DLL within .JAR
            } catch (IOException e1) {
                throw new RuntimeException(e1);
            }
        }
    }

    public static void main(String[] args) throws IOException, WriteException, InterruptedException {
        GUI gui = new GUI();
        gui.init();
    }
}
