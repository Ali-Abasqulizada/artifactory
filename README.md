# README #

Artifactory is a stateful service that'll host versioned static data such as model or airport databases (and other types of fiels) that we deploy on our production servers.

It runs on flask and uses sqlite as the database backend.

### How do I get set up? ###

#### Development environment setup
Create an sqlite database using the database schema from `db_schema.sql`.
With this command in Bash command line:

```
sqlite3 artifactory.db < db_schema.sql
```

This command will start a flask backend on port 5000. Access http://localhost:5000 in your browser
and confirm you are seeing the "Hello, World" message served by flask.


### Contribution guidelines ###

#### Committing work & Code Reviews

Main branch contains reviewed code ready for deployment. Committing directly to the main branch is not
allowed. Instead, create your own branch off of the main branch and use it for your development. When ready,
start a pull request to kick off the code review. Eventually, when the code review is completed, your development
branch will be merged with the main branch.



