/**
/* shortener.ts
/* URL短縮ツール「ゑびす短縮er」 
*/

// import global interface
import {} from "./@types/globalsql";
// モジュール読み込み
import { config as dotenv } from "dotenv"; // dotenv
import * as path from "path"; // path
import express from "express"; // express
import helmet from "helmet"; // セキュリティ対策
import crypto from "crypto"; // セキュリティ対策
import SQL from "./class/MySql0517"; // DB操作用
import Logger from "./class/Logger0516"; // ロガー

// ロガー
const logger: any = new Logger(__dirname, "");
// 開発フラグ
const DEV_FLG: boolean = false;

// モジュール設定
dotenv({ path: path.join(__dirname, ".env") });

// 開発環境切り替え
let globalDefaultPort: number; // ポート番号
let sqlHost: string; // SQLホスト名
let sqlUser: string; // SQLユーザ名
let sqlPass: string; // SQLパスワード
let sqlDb: string; // SQLデータベース名

// 開発モード
if (DEV_FLG) {
  globalDefaultPort = Number(process.env.DEV_PORT); // ポート番号
  sqlHost = process.env.SQL_DEVHOST!; // SQLホスト名
  sqlUser = process.env.SQL_DEVADMINUSER!; // SQLユーザ名
  sqlPass = process.env.SQL_DEVADMINPASS!; // SQLパスワード
  sqlDb = process.env.SQL_DEVDBNAME!; // SQLデータベース名
} else {
  globalDefaultPort = Number(process.env.DEFAULT_PORT); // ポート番号
  sqlHost = process.env.SQL_HOST!; // SQLホスト名
  sqlUser = process.env.SQL_ADMINUSER!; // SQLユーザ名
  sqlPass = process.env.SQL_ADMINPASS!; // SQLパスワード
  sqlDb = process.env.SQL_DBNAME!; // SQLデータベース名
}
// DB設定
const myDB: SQL = new SQL(
  sqlHost, // ホスト名
  sqlUser, // ユーザ名
  sqlPass, // ユーザパスワード
  Number(process.env.SQL_PORT), // ポート番号
  sqlDb // DB名
);
// express設定
const app: any = express(); // express

app.use(helmet()); // ヘルメットを使用する
app.set("view engine", "ejs"); // ejs使用
app.use(express.static("public")); // public設定
app.use(express.json()); // json設定
app.use(
  express.urlencoded({
    extended: true, // body parser使用
  })
);

// トップページ
app.get("/test", async (_: any, res: any) => {
  // ファイル名
  logger.debug("connented");
  res.send("connented");
});

// 短縮URLアクセス時
app.get("/:key", async (req: any, res: any) => {
  try {
    // 検索キー設定
    const searchKey: string = req.params.key;

    // 5文字なら処理
    if (searchKey.length == 8) {
      // 対象データ
      const shortUrlArgs: selectargs = {
        table: "shortenurl",
        columns: ["short_url", "usable"],
        values: [[searchKey], [1]],
        fields: ["pre_url"],
      };
      // 短縮前URL抽出
      const tmpPreUrlData: any = await myDB.selectDB(shortUrlArgs);

      // 該当URLにリダイレクト
      if (tmpPreUrlData != "error") {
        res.redirect(301, tmpPreUrlData[0].pre_url);
      } else {
        res.send("error");
      }
    } else {
      res.send("connected");
    }
  } catch (e: unknown) {
    // エラー型
    if (e instanceof Error) {
      logger.error(e.message);
    }
    res.send("error");
  }
});

// 短縮URL作成時
app.post("/create", async (req: any, res: any) => {
  try {
    // カウンタ
    let num: number = 0;
    // 検索結果
    let shortDataCount: number = 0;
    // 短縮文字列
    let randomKey: string = "";
    // 短縮対象URL
    const setUrl: any = req.body.url;

    // 重複無し生成までループ
    while (num < 5) {
      // 短縮文字列
      randomKey = createRandomString(8);
      // 対象データ
      const shortSelectArgs: countargs = {
        table: "shortenurl", // テーブル
        columns: ["short_url"], // カラム
        values: [[randomKey]], // 値
      };
      // 対象データ取得
      shortDataCount = await myDB.countDB(shortSelectArgs);

      // 検索結果あり
      if (shortDataCount == 0) {
        break;
      }
      // カウントアップ
      num++;

      // ループリミット超え
      if (num == 5) {
        throw new Error("randomkey making failed.");
      }
    }

    // 対象データ
    const insertTransArgs: insertargs = {
      table: "shortenurl",
      columns: ["pre_url", "short_url", "usable"],
      values: [setUrl, randomKey, 1],
    };
    // トランザクションDB格納
    const tmpReg: any = await myDB.insertDB(insertTransArgs);

    // エラー
    if (tmpReg == "error") {
      throw new Error("shortenurl insertion error");
    } else {
      logger.debug("initial insertion to shortenurl completed.");
      // 結果を返す
      res.send(randomKey);
    }
  } catch (e: unknown) {
    // エラー型
    if (e instanceof Error) {
      logger.error(e.message);
    }
    res.send("error");
  }
});

// ポート開放
app.listen(globalDefaultPort, () => {
  logger.debug(`shortener app listening at ebs.lol:${globalDefaultPort}`);
});

// 乱数生成器
const createRandomString = (length: number): string => {
  const S = "abcdefghijklmnopqrstuvwxyz";
  let buf = crypto.randomBytes(length);
  let rnd = "";
  for (var i = 0; i < length; i++) {
    rnd += S.charAt(Math.floor((buf[i] / 256) * S.length));
  }
  return rnd;
};
