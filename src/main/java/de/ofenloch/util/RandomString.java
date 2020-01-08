package de.ofenloch.util;

import java.nio.charset.Charset;
import java.util.Random;

public class RandomString {

    static final String AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + "0123456789" + "abcdefghijklmnopqrstuvxyz";
    static final String NonAlphaNumericString = "ÄÖÜäöüß" + AlphaNumericString;

    /**
     * generate a random alphanumeric string of length n
     * 
     * @param n length of the string
     * @return the string
     */
    static String getAlphaNumericString(final int n) {
        // length is bounded by 256 Character
        byte[] array = new byte[256];
        new Random().nextBytes(array);

        String randomString = new String(array, Charset.forName("UTF-8"));

        int charsLeft = n;
        // Create a StringBuffer to store the result
        StringBuffer rndString = new StringBuffer();
        // Append first n alphanumeric characters
        // from the generated random String into the result
        for (int k = 0; k < randomString.length(); k++) {

            char ch = randomString.charAt(k);

            if (((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9')) && (charsLeft > 0)) {
                rndString.append(ch);
                charsLeft--;
            }
        } // for (int k = 0; k < randomString.length(); k++)
          // return the resulting string
        return rndString.toString();
    } // static String getAlphaNumericString(final int n)

    /**
     * generate a random non-alphanumeric string of length n
     * 
     * @param n length of the string
     * @return the string
     */
    static String getNonAlphaNumericString(final int n) {
        // length is bounded by 256 Character
        byte[] array = new byte[256];
        new Random().nextBytes(array);

        String randomString = new String(array, Charset.forName("UTF-8"));

        int charsLeft = n;
        // Create a StringBuffer to store the result
        StringBuffer rndString = new StringBuffer();
        // Append first n alphanumeric characters
        // from the generated random String into the result
        for (int k = 0; k < randomString.length(); k++) {

            char ch = randomString.charAt(k);

            if ((ch >= ' ' && ch <= 'Ї') && (charsLeft > 0)) {
                rndString.append(ch);
                charsLeft--;
            }
        } // for (int k = 0; k < randomString.length(); k++)
          // return the resulting string
        return rndString.toString();
    } // static String getNonAlphaNumericString(final int n)

    public static void main(String[] args) {

        // Get the size n
        int n = 10;
        if (args.length > 0) {
            n = Integer.parseInt(args[0]);
        }

        // Get and display the alphanumeric string
        final String alphaNumStr = RandomString.getAlphaNumericString(n);
        final String nonAlphaNumStr = RandomString.getNonAlphaNumericString(n);
        System.out.println("alpha numeric random string:     \"" + alphaNumStr + "\"");
        System.out.println("    length : " + alphaNumStr.length() + "");
        System.out.println("non-alpha numeric random string: \"" + nonAlphaNumStr + "\"");
        System.out.println("    length : " + nonAlphaNumStr.length() + "");
    } // public static void main(String[] args)

} // public class RandomString