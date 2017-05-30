DROP TABLE IF EXISTS tweet CASCADE;
CREATE TABLE IF NOT EXISTS tweet (
	handle integer NOT NULL,
	timeTweeted timestamp NOT NULL,
	PRIMARY KEY(timeTweeted, handle),
	favouriteCount integer CONSTRAINT positive_favcount CHECK (favouriteCount >= 0),
	retweetCount integer CONSTRAINT positive_retweetcount CHECK (retweetCount >= 0),
	text varchar NOT NULL,
	originalAuthor varchar
);

DROP TABLE IF EXISTS hashtag CASCADE;
CREATE TABLE IF NOT EXISTS hashtag (
	tag varchar PRIMARY KEY
);

DROP TABLE IF EXISTS has CASCADE;
CREATE TABLE IF NOT EXISTS has (
	timeTweeted timestamp NOT NULL,
	handle integer NOT NULL,
	FOREIGN KEY (timeTweeted, handle)
		references tweet(timeTweeted, handle),
	tag varchar
		references hashtag(tag)
);
