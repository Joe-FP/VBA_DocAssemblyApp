## VBA Document Assembly Application

An extensive attempt at creating a desktop document assembly application for Microsoft Word, using Visual Basic for Applications (VBA). This application was designed to simplify the assembly & production process of a group of related documents, such as engineering work packages. All source files included. Word add-in applications are loaded as macro-enabled global templates (.dotm) into Word's Startup directory.

Created by Joseph Pollock, josephfpollock@gmail.com
For app demo, see: www.docu-mate.com/demo

 ### **For employers, see:**

**_standard modules/modGather.txt -- line 820 -- Private Function fcnGatherFrom(ByRef oDoc As Word.Document) As Variant_**

This function accepts a document object, and returns a sorted variant array containing all variable data which has been gathered from the document. The challenge here was creating a bespoke sort procedure for the array. Variable data should be returned in order of occurence, but header & footer variables should preceed all others.

**_standard modules/modGather.txt -- line 158 -- Public Function fcnInterferringOrAdjactentCC(Optional ByVal lngCCType As Long = 1) As Boolean_**

This function was a collaberation between myself and Microsoft Word MVP Gregory K. Maxey. Essentially, it stops document variables from being added inside, or directly next to other existing variables. Interferring variables cause a multitude of issues, and finding a reliable method of detecting interferring variables proved much more challenging that initially thought.

### **App summary:**

Firstly, documents are marked up using variables, which act as placeholders for document content. Ideally, the marked-up 
documents will now function as templates, in order to be reused. Next, we select the required documents for a new document 
package. The variables contained within the selected documents are then gathered and displayed together on the application 
interface. Finally, we select values for the variables depending on their type. All document variables are then populated, 
and the package is complete. Any type of content can be stored in the user library for reuse, making the application 
progressively more efficient.

### **Folder structure:**

- **standard modules**    - Variable gathering, markup, and Ribbon control. General support, and array support (by Chip Pearson).
                      Public variables, constants, and API declarations. Debugging support.

- **class modules**       - Global application event handlers, document event handlers, node and treeview code (by 
                      JKP Application Development Services (c)), dynamic runtime control event handlers.
                      
- **form modules**        - All userform source code for 7-no. userforms.



  
