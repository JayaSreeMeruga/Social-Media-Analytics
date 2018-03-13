cd /home/jkukkala/workspace/JAVA_BATCH_Code_Lenove/src/main/java/com/aail/JavaBatchcode
rm -f JavaBatchcode.jar
rm -rf classes
mkdir classes

export CLASSPATH=/home/jkukkala/workspace/JAVA_BATCH_Code_Lenove/src/main/java/com/aail/JavaBatchcode/jar_files/\*:./\*
javac  -d classes JavaBatchcode_Excel_Sheet.java
jar -cvf JavaBatchcode.jar -C classes/ . 

java com.aail.JavaBatchcode.JavaBatchcode_Excel_Sheet
