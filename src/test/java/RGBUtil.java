public class RGBUtil {
    public static void main(String[] args) {
        String r = Integer.toHexString(255).toUpperCase();
        String g = Integer.toHexString(255).toUpperCase();
        String b = Integer.toHexString(255).toUpperCase();

        System.out.println("#" + r+g+b);
    }
}
