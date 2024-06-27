<p align="right"><a href="https://github.com/ojudz08/AutomationProjects/tree/main">Back To Main Page</a></p>


<!-- PROJECT LOGO -->
<br />
<div align="center">
<h1 align="center">Parsing the Transaction Details of a Bank Statement</h1>
</div>


<!-- ABOUT PROJECT -->
### About Project

Parses the transaction details table of a bank statement. It iterates to all pages, parses the table and append the transaction details into the Summary file.


<table>
<tr>
   <th>Sale Date</th>
   <th>Post Date</th>
   <th>Transaction</th>
   <th>Amount</th>
   <th>Total</th>
   <th>Statement Date</th>
   <th>Due Date</th>
</tr>
<tr>
   <td>MM/DD/YY</td>
   <td>MM/DD/YY</td>
   <td>Transaction info</td>
   <td>PHP</td>
   <td>PHP</td>
   <td>MM/DD/YY</td>
   <td>MM/DD/YY</td>
</tr>
<tr>
   <td>MM/DD/YY</td>
   <td>MM/DD/YY</td>
   <td>Transaction info</td>
   <td>PHP</td>
   <td>PHP</td>
   <td>---</td>
   <td>---</td>
</tr>
<tr>
   <td>MM/DD/YY</td>
   <td>MM/DD/YY</td>
   <td>Transaction info</td>
   <td>PHP</td>
   <td>PHP</td>
   <td>---</td>
   <td>---</td>
</tr>
</table>


### What the output looks like

Below is a sample bankk statement for February 2024 showing only the transaction details.

   <img src="img/image1.png" alt="drawing" width="450"/>

___

Once you ran the sript, here's what the output looks like

   <img src="img/image2.png" alt="drawing" width="450"/>

___

### Requirements

This needs to be configured for the tabula-py to work in your system.

1. Download jdk 8 from the archive downloads (this [link](https://www.oracle.com/in/java/technologies/javase/javase8-archive-downloads.html) )
2. Install the jdk 8.
3. Set your JAVA_HOME to the path where your jdk 8 is installed. 
   ```
   C:\Program Files\Java\jdk-1.8
   ```
4. Add these to your environment variable
   ```
   C:\Program Files\Java\jdk-1.8
   C:\Program Files\Java\jdk-1.8\bin
   ```
5. Download the tabula-jar dependency from this [link](https://github.com/tabulapdf/tabula-java/releases). Then save it within the JAVA_HOME\jre\lib

6. Add the path to your environment variable
   ```
   JAVA_HOME\jre\lib\{tabula-jar-dependency}
   ```

___

### Installation

```Python version 3.11.9```

Note: **conda env** was used within VSCode to isolate the modules and dependencies used when creating this script. You may opt to create your conda env. Refer to this link [how to create conda env in VSCode.](https://code.visualstudio.com/docs/python/environments)

Run the command to install the libraries

```bat
python -m pip install -r requirements.txt
```



<!-- CONTACT -->
### Disclaimer

This project was created using Windows, the run.bat will only work with Windows. Please contact Ojelle Rogero - ojelle.rogero@gmail.com for any questions with email subject "Github Parsing PDFs".