import org.opencv.core.Mat;

/**
 * Created by Krzysiek on 2016-07-23.
 */
public interface VideoProcessor {
    Mat process(Mat inputImage);

    void setImageThreshold(double imageThreshold);

    void setHistory(int history);

}
