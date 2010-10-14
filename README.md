Crossfit Scorekeeping Generator
===============================

Crossfit competitions usually consist of multiple events.  Events are either time-based or work-based.  For time-based workouts, the lowest score wins, and for work-based workouts, the largest score (usually weight or reps) wins.  During the 2008 Crossfit Games, Crossfit HQ announced their scoring format as follows:  Every event is scored as usual, and each athlete is ranked among their peer group for that workout. This continues for each workout during the competition.  The athlete with the lowest total ranking wins.  (During the 2008 games, the better ranked you were the more rest you got between workouts as well.)

This is a great way to run multi-event competitions, but it can be a pain to keep scores and ranks in order.  While it would be fairly easy to write a web application to track all this, it's much easier to keep all the scores and rankings in full view, in something we're familiar with.

How this works
--------------

This script generates an Excel spreadsheet for multi-event competitions, using Excel formulas wherever possible.  All ranking is done for you within the spreadsheet.  You only need to enter three types of information:

 * Athlete Names
 * The ranking order of each event (0 = highest number wins, 1 = lowest number wins)
 * Scores for each athlete for each event

That's it!  The spreadsheet will rank every event and the overall competition every single time you update a score.  It's that easy.

TODO
----

This was a one-night hack, so it's not as smooth as it could be.  Here are some features I'd like to add:

 * Allow the bash script to accept arguments for number of athletes and number of events
 * Allow you to pass in the names of the events
 * Allow you to pass in the ranking order of the events, so you don't have to modify them in the spreadsheet
 * Protect the formula fields so you can't edit them.  I have no idea if this is even possible in Excel.
 * Release this as a gem so it's nicer to interact with

Also as a TODO is to continually hassle Crossfit HQ to release their "official" scorekeeping web application so affiliates can use it for their own competitions.  Open source fitness needs to upgrade their open source software!
 