DROP TABLE IF EXISTS users;
CREATE TABLE users (
  id integer primary key autoincrement,
  code string not null,
  name string not null,
  uuid string not null,
  rowIndex string not null
);
