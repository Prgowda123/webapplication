package extra;
 
import java.util.Random;

import org.testng.annotations.Test;
 
public class RandomStringGenerator {
	@Test
    public  void main() {
        int numberOfStrings = 100; // Number of random strings to generate
        int length = 6; // Length of each random string
 
        // Generate and print 100 random strings
        for (int i = 0; i < numberOfStrings; i++) {
            System.out.println(generateRandomString(length));
        }
    }
 
    public static String generateRandomString(int length) {
        StringBuilder randomString = new StringBuilder();
        Random random = new Random();
 
        // Ensure the first character is a letter
        char firstChar = (char) ('A' + random.nextInt(26)); // Random uppercase letter
        randomString.append(firstChar);
 
        // Generate the rest of the string with letters and digits
        for (int i = 1; i < length; i++) {
            int randType = random.nextInt(2); // Randomly decide if it's a letter or digit
            if (randType == 0) {
                // Random letter (uppercase)
                char randLetter = (char) ('a' + random.nextInt(26));
                randomString.append(randLetter);
            } else {
                // Random digit
                char randDigit = (char) ('0' + random.nextInt(10));
                randomString.append(randDigit);
            }
        }
 
        return randomString.toString();
    }
}