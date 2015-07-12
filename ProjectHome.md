# Roboshell is an advanced programming shell that makes it easy to rapidly deliver robust PC-based applications #


---


**Roboshell is used for developing PC-based programs that work with relational data.  Programs created using Roboshell have access to numerous built-in features and employ objects that make it easy to get common tasks done quickly, whilst applying the design principles you need. Robeshell makes development, deployment and maintenance more effective because the groundwork is done already.**

Roboshell is a productivity tool for experienced coders. It frees you from the boring details so you can focus on tailoring features your end Users want. Forget coding up the User Authentication, roles and form objects. It also helps insure against having to go back and redo basics that you never thought would become important later.

Applications commonly use static forms (i.e. windows) to present an interface to the user. Often, generating and modifying these 'views' is labourious and requires a lot of work to implement, test and deploy. Ongoing modifications/maintenance is typically where time 'begins to stand still'. Many applications use dynamic languages to enable the user interface to be 'painted' uniquely for each user. Typically, the processing required to do this within the application or in additional network tiers can be slow. In such environments transactions can be difficult to log well and error messages easily lost, making debugging troublesome.

Roboshell takes the best of both worlds- automatically generating each 'view' from logic and data in the database, where it is executed with safety and speed. Roboshell programs do not require a set of inter-dependent forms to be maintained, compiled into releases and rolled-out to each desktop. Roboshell applications do not even need to be installed on the client computer to run and can be easily implemented within enterprise domains and policy.

Programs built in Roboshell naturally excel in data entry, account and workflow maintenance. Wherever user-interface performance is vital, Roboshell is particularly well proven in supporting back-office tasks, accounting teams and customer facing staff.


**What Roboshell is not:**

  1. A point and click programming application: It does not aim to limit creativity or the features implemented in your applications.
  1. A Rapid Application Development 'Tool'. Many such tools tend to help programmers to build around key design aspects (such as keeping a clean object or authentication model). This tends to expedite early development, but require bug-fixing frenzies, and/or otherwise cripple the application's life and utility.  That said, Roboshell can be used to deliver programs very promptly indeed :)
  1. Bloatware: Roboshell applications do not have a large footprint, i.e. the number of lines of code compiled are minimal. Libraries and externally called applications are completely optional. The Environmental and networking pre-requisites required to run your program are simple too (see below).

**Project aims**

Roboshell aims to relieve programmers of the hard work needed to build and maintain quality applications throughout the software lifecycle. Once you use Roboshell to build a program, you see that the more labourious programming aspects are mostly done. Subsequently, the functionality you want to build, be it new functionality or just enhancements and maintenance, can be done very quickly.

**What is so good about it:**

  1. A clean and well-built security model
  1. User interface design that is highly workflow oriented;
    * display and management of data is fast (very)
    * ideal for workplace and process automation applications
  1. Easy deployment.  When you build a program in Roboshell, core logic is stored in the database, not statically compiled into the application where it can require a user to carry out updates. Most updates and functional enhancements can be executed with little or no effort on the client desktop; simply applying standard practices when updating the database ensures stability and a high degree of change assurance.
  1. No installaion required.  The client program can often be run from a remote fileshare.
  1. Roboshell programs uniquely dove-tail to their back-end database, providing a number of advantages not seen in other shells and development tools
  1. User actions are easily recorded in an enterprise-class transaction log
  1. Access to any object can be enabled and disabled on-the-fly in the application configuration, without any change to the application itself.

With Roboshell, it is easy to build in enterprise class authentication. Choose to control user-privilages however it is best for you. Access to all application configuration, objects and data within the security model can be set to use internal or external authentication services

Presently access is controlled by 2 methods:

  1. Database authentication
  1. MS-Windows based (NT) authentication, a.k.a. Active Directory
  1. Other- If you want to, a Roboshell program can use anything- e.g. configuration files or alternate authentication services.

Roboshell supports the development of all kinds of Programs from basic to complex.  It places no limitation or assumptions on what you choose to create, rather it gives you the flexibility to do what you want.


---


**Pre-requisites:**

Based on C# and MS-SQL, Roboshell requires only an SQL runtime or database server, Microsoft .Net framework 2.0+ (or mono).

  * Server host:
> MS-SQL Server 2005, 2008, Microsoft .Net 2.0 (or higher, or mono)

  * Client computer:
> Microsoft .Net 2.0 (or higher, or mono)

  * Developer workstation:
> MS-SQL Server 2005, 2008, SQL-Management Studio or Query Analyzer,
> Microsoft .Net 2.0 (or higher, or mono)

  * Network Requirements:
> The client must be able to connect to the back end database. Ports and authentication methods are all user configurable.
> Applciation design can extend these requirements depending on the features implemented.


**Present Audience:**

Individuals, Groups and Organisations using SQL server backends that need to build and maintain fast, reliable software. Ideal for Client-Server networks and programs distributed to end-users.

**Status:**
This project is in beta phase, and may remain in beta for 'quite some time'.  We do not claim it to be free of bugs. Use Roboshell at your own risk, and do not use it in production without proper testing. We make no undertakings about the possibility that your systems or data will not be adversely affected by Roboshell.  You should test (as no doubt you already do) to minimise this risk.

**Does it work?**
Yes, very well.

**Is it mature?**
Roboshell began on Sybase back in the 80s, so it is more mature than many :)

**Is it free of bugs?**
No. If you use it, you need to be tolerant of bugs. If you are not, you will need to help us squish 'em!

**Who uses it?**
To date, primarily orgnasiations in the Finance Industry.  During the development of Roboshell, numerous successful programs were built and remain very effectively maintained by small teams.  The Organisations using them have not been asked to acknowledge this publicly, so they are not listed.


---



**Why use Roboshell?**

You like the idea of
  1. a free ('as in beer') tool to help you write C# apps quickly.
  1. highly usable interfaces which require little support.
  1. not having to write all the 'pesky stuff' that is not core functionality.
  1. writing applications that are easy to maintain/adjust when necessary, with a minimum of complexity, testing and roll-out effort.
  1. providing the stability and reliability of programs written with an effective security model, without having to do a ground-up implementation each time to achieve it.
  1. Transaction Auditing: You want to monitor or debug what your app is doing.
  1. a simple and robust long-term application development construct, dependent on industry standard, proven technologies (such as MS-SQL and MS .Net).  Roboshell has no complicated dependencies and is very effective on MS-Windows platforms. Roboshell could be easily extended to support similar technologies on BSD, Linux, UNIX or Solaris if you were to want it to do so.
  1. Stability: Roboshell is a proven and advanced programming solution, but is not for novices.  You need to know how to program Stored Procedures, C#, RTM and ask questions to use it. Both Users and Contributors are welcome in the Roboshell project :)
  1. Longevity: You want the peace of mind that only 'having the sourcecode' brings; never having to hope a third party vendor's favour, focus or strategic direction will shift away from you.  Neither do you have to worry about them putting up their prices or changing their support service. See http://en.wikipedia.org/wiki/Apache_License for details.
  1. Security: Roboshell is inherrently secure as it implements user and group based security regimes based on what you have available to work with.  Roboshell supports Active Directory authentication as well (ie can be combined with) mixed-mode adn SQL-based authentication.  Both can be used simultaneously.
  1. Support: A community of developers is being encouraged and will in time, be able to provide  development support and maintenance assistance.
  1. Good context-sensitive documentation (ie what you write when building an applicaton): Roboshell contains an automated HTML help tool that generates CSS-driven help pages specific to the application's you build on it. The help features are built in from the start, so during development you can feed Roboshell your notes and other details. This is used to intelligently minimise the amount of documentation required to deliver effective, context friendly documentation.
  1. Tool documentation (i.e. for using Roboshell): We are writing this at present. A draft will be available soon.
  1. Roboshell is a stable and robust programming construct which actively fosters good programming principles, not dirty techniques.
  1. Roboshell is not based on transient technologies and will be just as useable and maintainable in 10 years as it is today.  SQL and C# are not going anywhere fast :)