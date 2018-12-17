FROM openjdk:8-jre-slim

# Add the jar with all the dependencies
# Maven creates TestSelKeyword1-0.0.1.jar.jar in the target folder of my workspace.
# We need this in some location - say - /usr/share/tag folder of the container
ADD  target/TestSelKeyword1-0.0.1.jar /usr/share/tag/TestSelKeyword1-0.0.1.jar

# Command line to execute the test
ENTRYPOINT ["/usr/bin/java", "-cp", "/usr/share/tag/TestSelKeyword1-0.0.1.jar", "org.testng.TestNG", "-testclass", "com.driver.MainTestNG"]