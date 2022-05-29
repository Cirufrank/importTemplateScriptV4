# Import Template Checker

![](https://media.giphy.com/media/aNqEFrYVnsS52/giphy.gif)

Are you a Data Specialist who's been tasked with checking import templates for clients? Wouldn't it be nice if a robot friend went ahead and did the ground work for you, leaving you with very little to check? Well look no further! I created this script to assist Galaxy Digital Data Specialists with checking clients' import templates, by automating the process.

**Link download add-on:** <<place_link_here>></br>
**Link documentation:** <<place_link_here>>

![alt tag](<<place_link_here>>)

## How It's Made:
I used Google's App Script language and API to write methods that perform the core template checks whenever triggred by the press of a button. I used OOP to organize my code by creating a general template class and additional template classes for each type of template that needs to be checked. Each class contains the properties and methods needed by the templates to perform their checks, and methods are shared between templates by utilizing superclasses and subclasses. 

## Optimizations
<ul>
  <li>De-duplicated code and made the checks more user-friendly by storing the potential titles of columns that need to be checked within an array, and having their check methods iterate through the potential titles within the array so that if multiple columns that need the same check are present, they are not missed, and this lets user not have to worry about varied columns titles. Additionally, through implementing this, I was able to remove multiple methods that performed checks based on a specific column title, and apply one method to all columns that contain certain header titles. </li>
  <li>After originally writing custom methods to perform the duplicate email check and white-space removal check, I re-wrote the methods to utilze Google Sheet's COUNTIF and trim functions. This improved the speed of the script and lowered the memory needed to run the script.</li>
</ul>


## Lessons Learned:
<ul>
  <li>Though writing custom methods for each task provides much creative freedom, it's important to keep the memory needed for the script, and the performance of the script in mind so that it is able to work effectivley for small data sets as well as large data sets.</li>
  <li>I learned of the importance of organization of code, and once my code was organized and OOP was fully implemented, I was allowed to seamlessly update methods, fix bugs, and add features to multiple checks at a time without having to duplciate methods.</li>
  <li>Additionally, user experience should aslo be kept at the forefront of the mind, and it's important to know when to put in some additional work in order to ensure the product can be used by others without headache.</li>
  <li>Lastly, at first I gave each sheet checked a custom name based on the type of check being ran (user template checks would get a sheet name of "User" and report sheet name of "User Report"). Though this was helpful, it would cause errors if a user needed to run multiple of the same type checks within the same spreadsheet (i.e. one user template has 50,000 rows so the user splits it up into 5 separate sheets, and runs checks on each of these sheets within the same spreadsheet). I ended up changing this, and keeping each original sheet name the same, and adding "Report" to the end of the sheet name for the reprot sheet. This enables users to run the same type of check for different sheets within the same spreadsheet.</li>
</ul>
    
