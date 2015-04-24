ALTER TABLE articles
ADD CONSTRAINT articles_pk PRIMARY KEY(article_id);


ALTER TABLE journals
ADD CONSTRAINT journals_pk PRIMARY KEY(journal_id);

ALTER TABLE articles
ADD CONSTRAINT articles_journals_fk FOREIGN KEY(journal_id) REFERENCES journals(journal_id)
  ON DELETE CASCADE
  ON UPDATE CASCADE;
  
ALTER TABLE authors
ADD CONSTRAINT authors_pk PRIMARY KEY(author_id);

ALTER TABLE articles_authors
ADD CONSTRAINT articles_authors_pk PRIMARY KEY(article_id, author_id);

ALTER TABLE articles_authors
ADD CONSTRAINT articles_authors_articles_fk FOREIGN KEY(article_id) REFERENCES articles(article_id)
  ON DELETE CASCADE
  ON UPDATE CASCADE;
  
ALTER TABLE articles_authors
ADD CONSTRAINT articles_authors_authors_fk FOREIGN KEY(author_id) REFERENCES authors(author_id)
  ON DELETE CASCADE
  ON UPDATE CASCADE;
