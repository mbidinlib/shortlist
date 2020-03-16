# shortlist
Stata program for first level shortlisting

# Overview

IPA Ghana conducts external recruitment of candidates for multiple roles. This stata program reads the downloaded csv application data from surveyCTO and conducts a first level shortlisting based on a certain criteria.


# installation(Beta)

net install shortlist, all replace ///
	from("https://raw.githubusercontent.com/mbidinlib/shortlist/edit/master/ado")

# Syntax

* gradeapplication using (full path ti criteria sheet)



