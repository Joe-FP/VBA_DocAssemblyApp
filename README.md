# VBA_DocAssemblyApp

An extensive attempt at creating a desktop document assembly application for Microsoft Word. All source files included. 
Word add-in applications are loaded as macro-enabled global templates (.dotm) into Word's Startup directory.

Created by Joseph F. Pollock, josephfpollock@gmail.com
For app demo, see: www.docu-mate.com/demo


App summary:

This application was designed to simplify the assembly & production of a group of related documents, such as engineering
work packages.

Firstly, documents are marked up using variables, which act as placeholders for document content. Ideally, the marked-up 
documents will now function as templates, in order to be reused. Next, we select the required documents for a new document 
package. The variables contained within the selected documents are then gathered and displayed together on the application 
interface. Finally, we select values for the variables depending on their type. All document variables are then populated, 
and the package is complete. Any type of content can be stored in the user library for reuse, making the application 
progressively more efficient.


Folder structure:

standard modules    - Variable gathering, markup, and Ribbon control. General support, and array support (by Chip Pearson).
                      Public variables, constants, and API declarations. Debugging support.

class modules       - Global application event handlers, document event handlers, node and treeview code (by 
                      JKP Application Development Services (c)), dynamic runtime control event handlers.
                      
form modules        - All userform source code for 7-no. userforms.


Highlights:

standard modules/modGather.txt - line 820 - Private Function fcnGatherFrom(ByRef oDoc As Word.Document) As Variant 
