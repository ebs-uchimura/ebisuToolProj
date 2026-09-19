/**
 * Sql.ts
 *
 * name：SQL
 * function：SQL operation
 * updated: 2024/05/17
 **/

// import global interface
import { } from "../@types/globalsql";

// define modules
import * as mysql from 'mysql2'; // mysql

// SQL class
class SQL {

  static pool: any; // sql pool
  static encryptkey: string; // encryptkey

  // construnctor
  constructor(host: string, user: string, pass: string, port: number, db: string, key?: string) {
    console.log('db: initialize mode');
    // DB config
    SQL.pool = mysql.createPool({
      host: host, // host
      user: user, // username
      password: pass, // password
      database: db, // db name
      port: port, // port number
      waitForConnections: true, // wait for conn
      idleTimeout: 1000000, // timeout(ms)
      insecureAuth: false // allow insecure
    });
    // encrypted key
    SQL.encryptkey = key!;
  }

  // inquire
  doInquiry = async (sql: string, inserts: string[]): Promise<any> => {
    return new Promise(async (resolve, reject) => {
      try {
        // make query
        const qry: any = mysql.format(sql, inserts);
        // connect ot mysql
        const promisePool: any = SQL.pool.promise(); // spread pool
        const [rows, _] = await promisePool.query(qry); // query name

        // empty
        if (SQL.isEmpty(rows)) {
          // return error
          throw new Error('error');

        } else {
          // result object
          resolve(rows);
        }

      } catch (e: unknown) {
        // エラー型
        if (e instanceof Error) {
          // error
          console.log(e.message);
          reject('error');
        }
      }
    });
  }

  // count db
  countDB = async (args: countargs): Promise<number> => {
    return new Promise(async (resolve) => {
      try {
        console.log('db: countDB mode');
        // total
        let total: number;
        // query string
        let queryString: string;
        // array
        let placeholder: any[];

        // query
        queryString = 'SELECT COUNT(*) FROM ??';
        // placeholder
        placeholder = [args.table];

        // if column not null
        if (args.columns.length > 0 && args.values.length > 0) {
          // add where phrase
          queryString += ' WHERE';
          // columns
          const columns: string[] = args.columns;
          // values
          let values: any[][] = args.values;
          // columns length
          const columnsLength: number = columns.length;

          // loop for array
          for (let i: number = 0; i < columnsLength; i++) {
            // add in phrase
            queryString += ' ?? IN (?)';
            // push column
            placeholder.push(columns[i]);
            // push value
            placeholder.push(values[i]);

            // other than last one
            if (i < columnsLength - 1) {
              // add 'and' phrase
              queryString += ' AND';
            }
          }
        }

        // do query
        await this.doInquiry(queryString, placeholder).then((result: any) => {
          // result exists
          if (result !== 'error') {
            total = result[0]['COUNT(*)'];

          } else {
            total = 0;
          }
          console.log(`count: total is ${total}`);
          // return total
          resolve(total);

        }).catch((err: unknown) => {
          // error
          console.log(err);
          resolve(0);
        });

      } catch (e: unknown) {
        // error
        console.log(e);
        resolve(0);
      }
    });
  }

  // select db
  selectDB = async (args: selectargs): Promise<any> => {
    return new Promise(async (resolve) => {
      try {
        console.log('db: selectDB mode');
        // query string
        let queryString: string;
        // array
        let placeholder: any[];

        // if fields exists
        if (args.fields) {
          // query
          queryString = 'SELECT ?? FROM ??';
          // placeholder
          placeholder = [args.fields, args.table];

        } else {
          // query
          queryString = 'SELECT * FROM ??';
          // placeholder
          placeholder = [args.table];
        }

        // if column not null
        if (args.columns.length > 0 && args.values.length > 0) {
          // add where phrase
          queryString += ' WHERE';
          // values
          let values: any[][] = args.values;
          // columns
          const columns: string[] = args.columns;

          // loop for array
          for (let i: number = 0; i < args.columns.length; i++) {
            // add in phrase
            queryString += ' ?? IN (?)';
            // push column
            placeholder.push(columns[i]);
            // push value
            placeholder.push(values[i]);

            // other than last one
            if (i < args.columns.length - 1) {
              // add 'and' phrase
              queryString += ' AND';
            }
          }
        }

        // if column not null
        if (args.spancol && args.span) {
          // query
          queryString += ' AND ?? > date(current_timestamp - interval ? day)';
          // push span column
          placeholder.push(args.spancol);
          // push span limit
          placeholder.push(args.span);
        }

        // query
        queryString += ' ORDER BY ??';

        // if reverse
        if (args.reverse) {
          // query
          queryString += ' ASC';

        } else {
          // query
          queryString += ' DESC';
        }

        // if order exists
        if (args.order) {
          // push order key
          placeholder.push(args.order);

        } else {
          // push default id
          placeholder.push('id');
        }

        // if limit exists
        if (args.limit) {
          // query
          queryString += ' LIMIT ?';
          // push limit
          placeholder.push(args.limit);
        }

        // if offset exists
        if (args.offset) {
          // query
          queryString += ' OFFSET ?';
          // push offset
          placeholder.push(args.offset);
        }

        // do query
        await this.doInquiry(queryString, placeholder).then((result2: any) => {
          resolve(result2);
          console.log('select: db select success');

        }).catch((err: unknown) => {
          // error
          console.log(err);
          resolve('error');
        });

      } catch (e: unknown) {
        // error
        console.log(e);
        resolve('error');
      }
    });
  }

  // update
  updateDB = async (args: updateargs): Promise<any> => {
    return new Promise(async (resolve1) => {
      try {
        console.log('db: updateDB mode');

        // プロミス
        const promises: Promise<any>[] = [];

        // ループ
        for (let i = 0; i < args.setcol.length; i++) {
          // プロミス追加
          promises.push(
            new Promise(async (resolve2, reject2) => {
              // query string
              let queryString: string = 'UPDATE ?? SET ?? = ? WHERE ?? = ?';
              // array
              let placeholder: any[] = [
                args.table,
                args.setcol[i],
                args.setval[i],
                args.selcol[i],
                args.selval[i],
              ];

              if (args.spancol && args.spanval) {
                queryString += ' AND ?? < date(current_timestamp - interval ? day)';
                placeholder.push(args.spancol);
                placeholder.push(args.spanval);
              }

              // do query
              await this.doInquiry(queryString, placeholder).then((result: any) => {
                resolve2(result);
                console.log('select: db update success');

              }).catch((err: unknown) => {
                // error
                console.log(err);
                reject2('error');
              });
            })
          )
        }
        // 全終了
        Promise.all(promises).then((results) => {
          resolve1(results);
        });

      } catch (e: unknown) {
        // error
        console.log(e);
        resolve1('error');
      }
    });
  }

  // insert
  insertDB = async (args: insertargs): Promise<any> => {
    return new Promise(async (resolve) => {
      try {
        console.log('db: insertDB mode');
        // columns
        const columns: string[] = args.columns;
        // values
        const values: any[] = args.values;
        // password index
        const passwordIdx: number = columns.indexOf('password');

        // include password
        if (passwordIdx > -1) {

          // it's string
          if (typeof (values[passwordIdx]) == 'string') {
            // password
            const passphrase: string = values[passwordIdx]

            // not empty
            if (passphrase != '') {
              // change to encrypted
              values[passwordIdx] = `HEX(AES_ENCRYPT(${passphrase},${SQL.encryptkey}))`;
            }
          }
        }
        // query string
        const queryString: string = 'INSERT INTO ??(??) VALUES (?)';
        // array
        const placeholder: any[] = [args.table, args.columns, values];

        // do query
        await this.doInquiry(queryString, placeholder).then((result: any) => {
          resolve(result);
          console.log('select: db insert success');

        }).catch((err: unknown) => {
          console.log(err);
          resolve('error');
        });

      } catch (e: unknown) {
        // error
        console.log(e);
        resolve('error');
      }
    });
  }

  // empty or not
  static isEmpty(obj: Object) {
    // check whether blank
    return !Object.keys(obj).length;
  }
}

// export module
export default SQL;