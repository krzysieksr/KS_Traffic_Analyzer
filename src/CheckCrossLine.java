import org.opencv.core.Point;
import org.opencv.core.Rect;

/**
 * Created by Krzysiek on 2016-07-26.
 */
public class CheckCrossLine {

    public int lAx;
    public int lAy;
    public int lBx;
    public int lBy;
    public double a;
    public double b;
    Point l1;
    Point l2;


    public CheckCrossLine(Point l1, Point l2) {
        this.l1 = l1;
        this.l2 = l2;
        this.lAx = (int) (Math.min(l1.x, l2.x));
        this.lAy = (int) (Math.min(l1.y, l2.y));
        this.lBx = (int) (Math.max(l1.x, l2.x));
        this.lBy = (int) (Math.max(l1.y, l2.y));
    }

    public boolean rectContainLine(Rect rect) {
        int PrA = (int) ((rect.tl().x + rect.br().x) / 2);      //straight "a"
        int PrB = (int) ((rect.tl().y + rect.br().y) / 2);      //straight "b"

//        int pktCx = PrA;
        int pktCy = (int) (rect.tl().y);

//        int pktDx = PrA;
        int pktDy = (int) (rect.br().y);

        int pktEx = (int) (rect.tl().x);
//        int pktEy = PrB;

        int pktFx = (int) (rect.br().x);
//        int pktFy = PrB;


        if (lBx != lAx && lBy != lAy) {

            this.a = (l2.y - l1.y) / (l2.x - l1.x);
            this.b = l1.y - a * l1.x;

            int crax = PrA;
            double cray = (a * PrA + b);

            double crbx = (PrB - b) / a;
            int crby = PrB;

            if ((lAx <= crax) && (lBx >= crax) &&
                    (lAy <= cray) && (lBy >= cray) &&
                    (pktCy <= cray) && (pktDy >= cray))
                return true;

            else if ((lAx <= crbx) && (lBx >= crbx) &&
                    (lAy <= crby) && (lBy >= crby) &&
                    (pktEx <= crbx) && (pktFx >= crbx))
                return true;

            else
                return false;
        } else
            return false;
    }
}