# Archived

This project is now archived and superseeded by the [rowlistlib2](https://gihub.com/francescofoti/rowlistlib2) project.

# Readme snapshot

rowlistlib
==========

A COM library implementing easy to use, searchable and sortable in memory multi-column lists and row objects, well suited to work with disconnected database recordsets.

What for ?
==========
When I'll add the database classes (just 2 or 3) and the database module, you'll see how this library can be used to access different database engines (MySQL, Access, SQL Server, ...) with the same code.
Although this type of use for this library, for "universal" database access, is only one possible application.

Also, this is written in VB, for reuse in VB5/6, all VBA capable Microsoft Office applications, 32 & 64 bits compatible, and a COM library (compiled with VB5, which I'll add soon) that can be reused with any COM capable client.
That should cover the products and platforms on which we may be interested to leverage the library productivity benefits.

Release notes
=============

GPL v2: the modules sources still don't reference the license; they're all released in GPL v2.

The VBA2013 subdirectory contains an Access 2013 database with the library testdriver's modules.

The Doc subdirectory contains *outdated* documentation; I will update that stuff later.

This is a 32 and 64 bits compatible library.

Open RowListLibVBATestDriver.accdb with Access, go to VBA's debug window (CTRL+G) and type "Main" (without quotes) and then ENTER.
NB: When the database opens, you'll have to authorize it for running macros (see the yellow warning bar that appears under the ribbon).
