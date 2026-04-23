//　define modules
const express = require('express'); //　express
const path = require('path'); // path
const mysql = require('mysql'); // mysql
const bodyParser = require('body-parser'); // body-parser
const cookieParser = require('cookie-parser'); // cookie-parser
const session = require('express-session'); // express-session
const fs = require('fs'); // fs
const crypto = require("crypto"); // crypto
const passport = require('passport'); // passport
const LocalStrategy = require('passport-local').Strategy; // passport strategy
const nodemailer = require('nodemailer'); // nodemailer

// enable URL encode 
var urlencodedParser = bodyParser.urlencoded({extended : true});

//　EXPRESS
var app = express();

//　NODEMAILER configuration ↓----------------------------------
//　mail initializing
var transporter = nodemailer.createTransport({
    host: '*****', //　host
    port: 25, // port number
    use_authentication: false, // not authorize
    tls: {
        rejectUnauthorized: false // do not fail on invalid certs
      }
    });
//　mail options
var mailOptions = {
    from: '*****', // sender
    to: '*****', // receiver
    subject: '*****',　// title
    text: ''　// body
  };

// PASSPORT configuration↓----------------------------------
// passport initialization
app.use(passport.initialize());
// session config
app.use(session({
    // secret word
    secret: "*****",
    // overwrite option
    resave: false,
    // save uninitialized
    saveUninitialized: false,
    // cookie configuration
    cookie: {
        // cookie only http use
        httpOnly: true, 
        // use secret
        secure: false, 
        // cookie keep time
        maxage: 1000 * 60 * 30
      }
    }));
// passport
app.use(passport.session());

// EXPRESS configuraiton ↓-----------------------------------------
//　view config
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');
// URL encoding
app.use(bodyParser.urlencoded({ limit:'50mb', extended: true }));
//app.use(express.urlencoded({ extended: true }));
// JSON parsing
app.use(bodyParser.json({limit: '50mb'}));
// asset root folder
app.use(express.static(path.join(__dirname, 'public')));

//　MYSQL configuration↓---------------------------------------
//　SQLserver config
const conn = mysql.createConnection({
  host: '*****', // host
  user: '*****', // user name
  password: '*****', // passport
  database: '*****', // DB name
  port: 3306 // port number
});

//　SQL connect
conn.connect((err) => {
  if (err) {
    // error
    console.log("e-101: SQL Error" + err);
  }
  // success
  console.log("903: SQL connected");
});

// passport configuration↓---------------------------------------
// possport authorization config
passport.use(new LocalStrategy({
		session: true // enable session
  },
  function(username, password, done) {
    // extract user data
    conn.query('SELECT * FROM ?? WHERE ?? = ?', ['admin_user', 'login_id', username], function (err, user) {
      if (err) { 
        // DB error
        return done(err);
      }
      if (!user[0].login_id) {
        // authorization error
        return done(null, false, {message: 'Incorrect username.'});
      }
      // MD5 password
      var passwd = md5hex(password);
      // not correct password
      if (user[0].password !== passwd) {
        return done(null, false, {message: 'Incorrect password.'});
      }
      // password is correct
      return done(null, {
        id: user[0].login_id
      })
    });
  }
  ));
// serialize
passport.serializeUser(function(user, done) {
  done(null, user);
});
// deserialize
passport.deserializeUser(function(user, done) {
  done(null, user);
});

// variables ↓----------------------------------------------------
// global variable
var global_id = "";

// init data
init();

// management ↓---------------------------------------------------
// ★ ↓GET
// login
app.get('/login', function(req, res){
  res.render('login.ejs', {});
});

// edit
app.get('/edit_open', function(req, res){
  // confirm session
  if(req.session.passport){
    // get parameter
    // row number
    var index = Number(req.query.idx);
    // category
    var category = Number(req.query.rw);
    // order
    if(category == 1){
      // select order
      conn.query('SELECT * FROM ?? WHERE ?? = ?', ['order', 'id', index], function (err1, rows1) {
        if(err1){
          // SELECT error
          console.log("e-009: DB Error -SELECT- -man_order- : " + err);
        }
        if(rows1[0]){
          // modify time
          if(rows1[0].ordertime !== "0000-00-00 00:00:00"){
            var time = replaceDate(rows1[0].ordertime);
          }
          if(rows1[0].created_at !== "0000-00-00 00:00:00"){
            // created time
            var crtime = replaceDate(rows1[0].created_at);
          }
          if(rows1[0].updated_at !== "0000-00-00 00:00:00"){
            // updated time
            var uptime = replaceDate(rows1[0].updated_at);
          }
          // render page
          res.render('edit_pop.ejs', {title: "ORDER EDIT", idx: index, ctg: category, data: rows1[0], time: time, crtime: crtime, uptime: uptime});
        } else {
          // result is empty
          console.log("e-010: order table is empty");
        }
      });
      // user edit
    } else if(category == 2){
      // select user
      conn.query('SELECT * FROM ?? WHERE ?? = ?', ['user', 'id', index], function (err2, rows2) {
        if(err2){
          //　SELECT error
          console.log("e-009: DB Error -SELECT- -man_order- : " + err2);
        }
        if(rows2[0]){
          // modify time
          if(rows2[0].created_at !== "0000-00-00 00:00:00"){
            // created time
            var crtime = replaceDate(rows2[0].created_at);
          }
          if(rows2[0].updated_at !== "0000-00-00 00:00:00"){
            // updated time
            var uptime = replaceDate(rows2[0].updated_at);
          }
          // page rendering
          res.render('edit_pop.ejs', {title: "USER EDIT", idx: index, ctg: category, data: rows2[0], crtime: crtime, uptime: uptime});
        } else {
          // result is empty
          console.log("e-010: order table is empty");
        }
      });
      // product edit
    } else if(category == 3){
      // select item
      conn.query('SELECT * FROM ?? WHERE ?? = ?', ['item', 'id', index], function (err3, rows3) {
        if(err3){
          // SELECT error
          console.log("e-009: DB Error -SELECT- -man_order- : " + err3);
        }
        if(rows3[0]){
          // modify time
          if(rows3[0].created_at !== "0000-00-00 00:00:00"){
            // created time
            var crtime = replaceDate(rows3[0].created_at);
          }
          if(rows3[0].updated_at !== "0000-00-00 00:00:00"){
            // updated time
            var uptime = replaceDate(rows3[0].updated_at);
          }
          // render page
          res.render('edit_pop.ejs', {title: "ITEM EDIT", idx: index, ctg: category, data: rows3[0], crtime: crtime, uptime: uptime});
        } else {
          // result is empty
          console.log("e-010: order table is empty");
        }
      });
      // EDIT AGENT
    } else if(category == 4){
      // select agent
      conn.query('SELECT * FROM ?? WHERE ?? = ?', ['agent', 'id', index], function (err4, rows4) {
        if(err4){
          // SELECT error
          console.log("e-009: DB Error -SELECT- -man_order- : " + err4);
        }
        if(rows4[0]){
          // modify time 
          if(rows4[0].created_at !== "0000-00-00 00:00:00"){
            // created time
            var crtime = replaceDate(rows4[0].created_at);
          }
          if(rows4[0].updated_at !== "0000-00-00 00:00:00"){
            // updated time
            var uptime = replaceDate(rows4[0].updated_at);
          }
          // render page
          res.render('edit_pop.ejs', {title: "AGENT EDIT", idx: index, ctg: category, data: rows4[0], crtime: crtime, uptime: uptime});
        } else {
          // result is empty
          console.log("e-010: order table is empty");
        }
      });
    } else if(category == 5){
      // select label
      conn.query('SELECT * FROM ?? WHERE ?? = ?', ['label', 'id', index], function (err5, rows5) {
        if(err5){
          // SELECT error
          console.log("e-009: DB Error -SELECT- -man_order- : " + err5);
        }
        if(rows5[0]){
          // modify time 
          if(rows5[0].created_at !== "0000-00-00 00:00:00"){
            // created time
            var crtime = replaceDate(rows5[0].created_at);
          }
          if(rows5[0].updated_at !== "0000-00-00 00:00:00"){
            // updated time
            var uptime = replaceDate(rows5[0].updated_at);
          }
          // render page
          res.render('edit_pop.ejs', {title: "EDIT LABEL", idx: index, ctg: category, data: rows5[0], crtime: crtime, uptime: uptime});
        } else {
          // result is empty
          console.log("e-010: order table is empty");
        }
      });
    } else {
      // select owner
      conn.query('SELECT * FROM ?? WHERE ?? = ?', ['owner', 'id', index], function (err6, rows6) {
        if(err6){
          // SELECT error
          console.log("e-009: DB Error -SELECT- -man_order- : " + err6);
        }
        if(rows6[0]){
          // modify time
          if(rows6[0].created_at !== "0000-00-00 00:00:00"){
            // created time
            var crtime = replaceDate(rows6[0].created_at);
          }
          if(rows6[0].updated_at !== "0000-00-00 00:00:00"){
            // updated time
            var uptime = replaceDate(rows6[0].updated_at);
          }
          // render page
          res.render('edit_pop.ejs', {title: "EDIT OWNER", idx: index, ctg: category, data: rows6[0], crtime: crtime, uptime: uptime});
        } else {
          // result is empty
          console.log("e-010: order table is empty");
        }
      });
    }
  } else {
    console.log("e-202: Session Error -SELECT- -man_edit_open- : ");
    // session time out
    res.render('custom_timeout.ejs', {});
  }
});

// delete
app.get('/del_open', function(req, res){
  // confirm session
  if(req.session.passport){
    // get parameter
    // row no
    var index = Number(req.query.idx);
    // category no
    var category = Number(req.query.rw);
    // category is 1
    if(category == 1){
      // select order
      conn.query('DELETE * FROM ?? WHERE ?? = ?', ['order', 'id', index], function (err, results) {
        if(err){
          // DELETE error
          console.log("e-042: DB Error -DELETE- -man_order- : " + err);
        }
        if(results){
          // delete succeed
          console.log("224: delete_success");
        }
      });
    // category is 2
  } else if(category == 2){
      // select user
      conn.query('DELETE * FROM ?? WHERE ?? = ?', ['user', 'id', index], function (err2, results2) {
        if(err2){
          // DELETE error
          console.log("e-042: DB Error -DELETE- -man_order- : " + err2);
        }
        if(results2){
          // delete succeed
          console.log("224: delete_success");
        }
      });
    // category is 3
  } else if(category == 3){
      // select item
      conn.query('DELETE * FROM ?? WHERE ?? = ?', ['item', 'id', index], function (err3, results3) {
        if(err3){
          // DELETE error
          console.log("e-042: DB Error -DELETE- -man_order- : " + err3);
        }
        if(results3){
          // delete succeed
          console.log("224: delete_success");
        }
      });
    // category is 4
  } else if(category == 4){
      // select agent
      conn.query('DELETE * FROM ?? WHERE ?? = ?', ['agent', 'id', index], function (err4, results4) {
        if(err4){
          // DELETE error
          console.log("e-042: DB Error -DELETE- -man_order- : " + err4);
        }
        if(results4){
          // delete succeed
          console.log("224: delete_success");
        }
      });
    // category is 5
  } else if(category == 5){
      // select label
      conn.query('DELETE * FROM ?? WHERE ?? = ?', ['label', 'id', index], function (err5, results5) {
        if(err5){
          // DELETE error
          console.log("e-042: DB Error -DELETE- -man_order- : " + err5);
        }
        if(results5){
          // delete succeed
          console.log("224: delete_success");
        }
      });
    } else {
      // select owner
      conn.query('DELETE * FROM ?? WHERE ?? = ?', ['owner', 'id', index], function (err6, results6) {
        if(err6){
          // DELETE error
          console.log("e-042: DB Error -DELETE- -man_order- : " + err6);
        }
        if(results6){
          // delete succeed
          console.log("224: delete_success");
        }
      });
    }
  } else {
    console.log("e-202: Session Error -SELECT- -man_edit_open- : ");
    // session timeout
    res.render('custom_timeout.ejs', {});
  }
});

// ★ ↓POST
// login authorization
app.post('/login', urlencodedParser, passport.authenticate('local'), function(req, res){
  // session authorization
  if(req.session.passport){
    res.render('manage.ejs', {});
  } else {
    // authorization error
    res.render('custom_unauthorized.ejs', {});
  }
});

// viewer
app.post('/view', function(req, res){
  // confirm session
  if(req.session.passport){
    // select order
    conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['order', 'id'], function (err, rows) {
      if(err){
        // SELECT error
        console.log("e-009: DB Error -SELECT- -man_order- : " + err);
      }
      if(rows){
        // modify times
        var times = [];
        rows.forEach(function(value){
          // time exists
          if(value.ordertime !== "0000-00-00 00:00:00"){
            // add time
            times.push(replaceDate(value.ordertime));
          }
        });
        console.log("207: man_order_search_success_view");
        // extract user
        conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['user', 'id'], function (err2, rows2) {
          if(err2){
            // SELECT error
            console.log("e-011: DB Error -SELECT- -man_user- : " + err2);
          }
          if(rows2){
            console.log("208: man_user_search_success_view");
            // extract item
            conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['item', 'id'], function (err3, rows3) {
              if(err3){
                // SELECT error
                console.log("e-012: DB Error -SELECT- -man_product- : " + err3);
              }
              if(rows3){
                console.log("209: man_item_search_success_view");
                // extract agent
                conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['agent', 'id'], function (err4, rows4) {
                  if(err4){
                    // SELECT error
                    console.log("e-014: DB Error -SELECT- -man_agent- : " + err4);
                  }
                  if(rows4){
                    console.log("210: man_agent_search_success_view");
                    // extract label
                    conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['label', 'id'], function (err5, rows5) {
                      if(err5){
                        // SELECT error
                        console.log("e-016: DB Error -SELECT- -man_label- : " + err5);
                      }
                      if(rows5){
                        console.log("211: man_label_search_success_view");
                        // extract owner
                        conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['owner', 'id'], function (err6, rows6) {
                          if(err6){
                            // SELECT error
                            console.log("e-018: DB Error -SELECT- -man_owner- : " + err6);
                          }
                          if(rows6){
                            console.log("212: man_owner_search_success_view");
                            // counteer decrement
                            var counter = -1;
                            // goto management page
                            res.render('manage_view.ejs', {data: {order: rows, user: rows2, item: rows3, agent: rows4, label: rows5, owner: rows6, times: times}});
                          } else {
                            // result is empty
                            console.log("e-020: owner table is empty");
                          }
                        });
                      } else {
                        // result is empty
                        console.log("e-019: label table is empty");
                      }
                    });
                  } else {
                    // result is empty
                    console.log("e-017: agent table is empty");
                  }
                });
              } else{
                // result is empty
                console.log("e-015: item table is empty");
              }
            });
          } else {
            // result is empty
            console.log("e-013: user table is empty");
          }
        });
      } else {
        // result is empty
        console.log("e-010: order table is empty");
      }
    });
  } else {
    // result is empty
    console.log("e-203: Session Error -SELECT- -man_view- : ");
    // session timeout
    res.render('custom_timeout.ejs', {});
  }
});

// edit 
app.post('/edit', function(req, res){
  // confirm session
  if(req.session.passport){
    // extract order
    conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['order', 'id'], function (err, rows) {
      if(err){
        // SELECT error
        console.log("e-009: DB Error -SELECT- -man_order- : " + err);
      }
      if(rows){
        // modify time
        var times = [];
        rows.forEach(function(value){
          // time exists
          if(value.ordertime !== "0000-00-00 00:00:00"){
            // add time
            times.push(replaceDate(value.ordertime));
          }
        });
        console.log("218: man_order_search_success_edit");
        // extract user
        conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['user', 'id'], function (err2, rows2) {
          if(err2){
            // SELECT error
            console.log("e-011: DB Error -SELECT- -man_user- : " + err2);
          }
          if(rows2){
            console.log("219: man_user_search_success_edit");
            // extract item
            conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['item', 'id'], function (err3, rows3) {
              if(err3){
                // SELECT error
                console.log("e-012: DB Error -SELECT- -man_product- : " + err3);
              }
              if(rows3){
                console.log("220: man_item_search_success_edit");
                // extract agent
                conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['agent', 'id'], function (err4, rows4) {
                  if(err4){
                    // SELECT error
                    console.log("e-014: DB Error -SELECT- -man_agent- : " + err4);
                  }
                  if(rows4){
                    console.log("221: man_agent_search_success_edit");
                    // extract label
                    conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['label', 'id'], function (err5, rows5) {
                      if(err5){
                        // SELECT error
                        console.log("e-016: DB Error -SELECT- -man_label- : " + err5);
                      }
                      if(rows5){
                        console.log("222: man_label_search_success_edit");
                        // extract owner
                        conn.query('SELECT * FROM ?? ORDER BY ?? DESC', ['owner', 'id'], function (err6, rows6) {
                          if(err6){
                            // SELECT error
                            console.log("e-018: DB Error -SELECT- -man_owner- : " + err6);
                          }
                          if(rows6){
                            console.log("223: man_owner_search_success_edit");
                            // decrement
                            var counter = -1;
                            // goto edit page
                            res.render('manage_edit.ejs', {data: {order: rows, user: rows2, item: rows3, agent: rows4, label: rows5, owner: rows6, times: times}});
                          } else {
                            // result is empty
                            console.log("e-026: owner table is empty");
                          }
                        });
                      } else {
                        // result is empty
                        console.log("e-025: label table is empty");
                      }
                    });
                  } else {
                    // result is empty
                    console.log("e-024: agent table is empty");
                  }
                });
              } else　{
                // result is empty
                console.log("e-023: item table is empty");
              }
            });
          } else {
            // result is empty
            console.log("e-022: user table is empty");
          }
        });
      } else {
        // result is empty
        console.log("e-021: order table is empty");
      }
    });
  } else {
    // result is empty
    console.log("e-204: Session Error -SELECT- -man_edit- : ");
    // session timeout
    res.render('custom_timeout.ejs', {});
  }
});

// edit order
app.post('/edit_data1', function(req, res){
	// now time
	var now = getTime();
	// confime session
  if(req.session.passport){
  	// update order
  	var prd_price = Number(req.body.price);
		// update price
  	var old_itemcount = Number(req.body.old_itemcount);　// buyed number before update
  	var new_itemcount = Number(req.body.itemcount); // buyed number after update
  	var delta_price = prd_price * (new_itemcount - old_itemcount); // price diff
  	var old_gloss = Number(req.body.gross);　// total price before update
  	var new_gross = old_gloss + delta_price;　// total price after update
  	// update order table
  	conn.query('UPDATE ?? SET ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ? WHERE ?? = ?', ['order', 'total', new_gross, 'date', req.body.date, 'ownerid', req.body.ownerid, 'paid', Number(req.body.paid), 'progress', Number(req.body.progress), 'updated_at', now, 'rand_id', req.body.rand], function (err, results) {
  		if(err){
        // UPDATE error
        console.log("e-038: DB Error -UPDATE- -man_order- : " + err);
      }
      if(results){
        // UPDATE success
        console.log("224: man_order_update_success : " + results);
        conn.query('UPDATE ?? SET ?? = ?, ?? = ?, ?? = ?, ??= ?, ?? = ?, ?? = ?, ?? = ? WHERE ?? = ?', ['order', 'total', new_gross, 'itemcount', req.body.itemcount, 'date', req.body.date, 'ownerid', req.body.ownerid, 'paid', Number(req.body.paid), 'progress', Number(req.body.progress), 'updated_at', now, 'id', req.body.id], function (err2, results2) {
         if(err2){
            // UPDATE error
            console.log("e-038: DB Error -UPDATE- -man_order- : " + err2);
          } else {
            // UPDATE success
            console.log("224: man_order_update_success : " + results2);
          }
        });
      }
    });
  } else {
    console.log("e-204: Session Error -SELECT- -man_edit- : ");
    // session timeout
    res.render('custom_timeout.ejs', {});
  }
});

// user edit
app.post('/edit_data2', function(req, res){
  // now time
  var now = getTime();
	// confirm session
  if(req.session.passport){
  	// update user
  	conn.query('UPDATE ?? SET ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ? WHERE ?? = ?', ['user', 'name', req.body.name, 'ruby', req.body.ruby, 'group', req.body.group, 'zip', req.body.zip, 'address', req.body.addr, 'tel', req.body.tel, 'mail', req.body.mail, 'ownerid', req.body.ownerid, 'updated_at', now, 'id', req.body.id], function (err3, results3) {
  		if(err3){
        // UPDATE error
        console.log("e-038: DB Error -UPDATE- -man_order- : " + err3);
      }
      if(results3){
        //　UPDATE success
        console.log("224: man_order_update_success : " + results3);
      }
    });
  } else {
    console.log("e-204: Session Error -SELECT- -man_edit- : ");
    // session time out
    res.render('custom_timeout.ejs', {});
  }
});

// item edit
app.post('/edit_data3', function(req, res){
  // now time
  var now = getTime();
	// confirm session
  if(req.session.passport){
    // update user
    conn.query('UPDATE ?? SET ?? = ?, ?? = ?, ?? = ?, ?? = ? WHERE ?? = ?', ['item', 'ganre', req.body.ganre, 'price', req.body.price, 'specific', req.body.specific, 'updated_at', now, 'id', req.body.id], function (err4, results4) {
      if(err4){
        //　UPDATE error
        console.log("e-038: DB Error -UPDATE- -man_order- : " + err4);
      }
      if(results4){
        //　UPDATE success
        console.log("224: man_order_update_success : " + results4);
      }
    });
  } else {
    console.log("e-204: Session Error -SELECT- -man_edit- : ");
    // session time out
    res.render('custom_timeout.ejs', {});
  }
});

// edit agent
app.post('/edit_data4', function(req, res){
  // now time
  var now = getTime();
  // confirm session
  if(req.session.passport){
    // update agent
    conn.query('UPDATE ?? SET ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ? WHERE ?? = ?', ['agent', 'ganre', req.body.ganre, 'name', req.body.name, 'zip', req.body.zip, 'addr', req.body.addr, 'tel', req.body.tel, 'mail', req.body.mail, 'ownerid', req.body.ownerid, 'updated_at', now, 'id', req.body.id], function (err5, results5) {
      if(err5){
        // UPDATE error
        console.log("e-038: DB Error -UPDATE- -man_order- : " + err5);
      }
      if(results5){
        // UPDATE success
        console.log("224: man_order_update_success : " + results5);
      }
    });
  } else {
    console.log("e-204: Session Error -SELECT- -man_edit- : ");
    // session timeout
    res.redirect('/login');
  }
});

// edit label
app.post('/edit_data5', function(req, res){
  // now time
  var now = getTime();
  // confirm session
  if(req.session.passport){
    // update label
    conn.query('UPDATE ?? SET ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ? WHERE ?? = ?', ['label', 'ganre', req.body.ganre, 'url', req.body.url, 'name', req.body.name, 'itemid', req.body.itemid, 'orderid', req.body.orderid, 'ownerid', req.body.ownerid,　'comment', req.body.comment, 'updated_at', now, 'id', req.body.id], function (err6, results6) {
      if(err6){
        // UPDATE error
        console.log("e-038: DB Error -UPDATE- -man_order- : " + err6);
      }
      if(results6){
        // UPDATE success
        console.log("224: man_order_update_success : " + results6);
      }
    });
  } else {
    console.log("e-204: Session Error -SELECT- -man_edit- : ");
    // session time out
    res.render('custom_timeout.ejs', {});
  }
});

// edit charge
app.post('/edit_data6', function(req, res){
  // now time
  var now = getTime();
	// confirm session
  if(req.session.passport){
    // update charge
    conn.query('UPDATE ?? SET ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ?, ?? = ? WHERE ?? = ?', ['owner', 'name', req.body.name, 'employeeid', req.body.employeeid, 'name', req.body.name, 'belonging', req.body.belonging, 'permission', req.body.permission, 'updated_at', now, 'id', req.body.id], function (err7, results7) {
      if(err7){
        // UPDATE error
        console.log("e-038: DB Error -UPDATE- -man_order- : " + err7);
      }
      if(results7){
        //　UPDATE success
        console.log("224: man_order_update_success : " + results7);
      }
    });
  } else {
    console.log("e-204: Session Error -SELECT- -man_edit- : ");
    // session time out
    res.render('custom_timeout.ejs', {});
  }
});

// general operation ↓---------------------------------------------------
// ★↓GET
// redirect
app.get('/', function(req, res){
  res.redirect('/kosei');
});

// thanks page
app.get('/comp', function(req, res){
  res.render('comp.ejs', {});
});

// home
app.get('/kosei', function(req, res, next){
	console.log("001: index");
  // goto top page
  res.render('index.ejs', {uid: "", num: "", flg: false});
});

// ★↓POST
// authorization
app.post('/auth', function(req, res, next){
	console.log("101: auth");
  // globalize user ID
  global_id = req.body.data.id;
  // extract agent
  conn.query('SELECT * FROM ?? WHERE ?? = ?', ['item', 'agentid', 1], function (err, rows) {
    if(err){
      //　SELECT error
      console.log("e-009: DB Error -SELECT- -man_order- : " + err);
    }
    if(rows){
      // return item info
      res.send(rows);
    }
  });
});

// sort label
app.post('/sort', function(req, res, next){
  console.log("801: sort");
  var lb_no = req.body.data.lb_no;
  if(lb_no == 0){
    // extract label
    conn.query('SELECT * FROM ??', ['label'], function (err, rows) {
      if(err){
        // SELECT error
        console.log("e-009: DB Error -SELECT- -man_order- : " + err);
      }
      if(rows){
        // return labels
        res.send(rows);
      }
    });
  } else {
    // extract label
    conn.query('SELECT * FROM ?? WHERE ?? = ?', ['label', 'ganre', lb_no], function (err, rows) {
      if(err){
        //　SELECT error
        console.log("e-009: DB Error -SELECT- -man_order- : " + err);
      }
      if(rows){
        // return labels
        res.send(rows);
      }
    });
  }
  
});

// register info
app.post('/post_json', function(req, res, next){
	// userID
	var user_id = req.body.data.uid;
  // orderID
  var order_id = user_id.substr(0,10);
	// order Info
	var order_data = req.body.data.sent_item.order;
	// total price
	var gross_price = req.body.data.sent_item.total;
	// user info
	var user_data = req.body.data.sent_item.user;
  // user env info
  var env_data = req.body.data.sent_item.environment;
	// font info
	var font_data = req.body.data.sent_item.font;
  // text info
  var text_data = req.body.data.sent_item.txt;
  var lb_id = 0;
  // payment info
  var paid_data = (req.body.paid);
  if(paid_data){
  	paid_data = 1; // init payment info
  }
  // progress info
  var prg_data = (req.body.progress);
  if(prg_data){
  	prg_data = 1; // init progress info
  }
  // image URL
  var img_urls = [];
  // now time
  var now = getTime();
	// get management no
	if(order_data || user_data || environment_data || font_data){
		console.log("009: data_ready");
		var lb_man = 0;
		console.log(order_data.length);
    for(var i=0;i<order_data.length;i++){
      if(order_data[i].lb_flg){
        lb_id = 0;
      } else {
        lb_id = Number(order_data[i].label_id) -1;
      }
      // add image URL
      img_urls.push("./uploads/" + user_id + "/resized_" + String(i+1) + ".png");
      // insert order info
      conn.query('INSERT INTO ?? (??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', ['order', 'itemid', 'order_id', 'itemcount', 'total', 'ordertime', 'ownerid', 'rand_id', 'labelid', 'price', 'itemname', 'username', 'userruby', 'usergroup', 'userzip', 'useraddress', 'usertel', 'usermail', 'date', 'url', 'paid', 'progress', 'agentid', 'agentname', 'main_txt', 'sub_txt', 'other_txt', 'created_at', 'updated_at', order_data[i].product_id, order_id, order_data[i].amount, gross_price, now, 1, user_id, lb_id, order_data[i].price, order_data[i].name, user_data.uname, user_data.ruby, user_data.group, user_data.zip, user_data.address, user_data.tel, user_data.mail, user_data.date, img_urls[i], paid_data, prg_data, 1, '*****', order_data[i].main_txt, order_data[i].sub_txt, order_data[i].other_txt, now, now], function (err, results) {
        // error
        if(err){
          console.log("e-003: DB Error -INSERT- -order- : " + err);
        }
        if(results){
          // when success
          console.log("205: order_insert_success " + i);
        } 
      });
    }
      // insert user info
      conn.query('INSERT INTO ?? (??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', ['user', 'name', 'ruby', 'group', 'address', 'tel', 'mail', 'date', 'rand_id', 'orderid', 'created_at', 'updated_at', user_data.uname, user_data.ruby, user_data.meigi, user_data.address, user_data.tel, user_data.mail, user_data.date, user_id, order_id, now, now], function (err2, results2) {
          // error
          if(err2){
            console.log("e-002: DB Error -INSERT- -user- : " + err2);
          }
          if(results2){
            // success
            console.log("206: user_insert_success");
          }
        });
    }　else {
        // data error
        console.log("e-007: DB preparation Error -man_insert-");
      }
    });

// redirect to payment page
app.post('/', function(req, res, next){
  // redirect to top page
  res.redirect('/kosei');
});

// payment page
app.post('/subpay', function(req, res, next){
  // now time
  var now = getTime();
  // got ID
  var b_id = req.body.data.id;
  // got data
  var b_d = req.body.data.data;
  // insert payment data
  conn.query('INSERT INTO ?? (??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', ['payment', 'rand_id', 'shop_id', 'order_id', 'password', 'shop_pass_string', 'ret_url', 'price', 'use_credit', 'datetime', 'job_cd', 'cancel_url', 'created_at', 'updated_at', b_id, b_d.ShopID, b_d.OrderID, b_d.Password, b_d.ShopPassString, b_d.RetURL, b_d.Amount, b_d.UseCredit, b_d.DateTime, b_d.JobCd, b_d.CancelURL, now, now], function (err, results) {
    if(err){
      // INSERT error
      console.log(err);
    }
    if(results){
      // INSERT success
      console.log("success");
    }   
  });
});

// img upload
app.post('/upload', function(req, res, next){
	console.log("006: uploaded")
  // init variable
  var got_arr = "";
  // get counter
  var counter = req.get('X-Custom-Header');
  // input array
  got_arr = counter.split(',');
	// image data
	var item = req.body.data.sent_item;
	// encoding data
  const encodedData = item;
  // file data
  const fileData = encodedData.replace(/^data:\w+\/\w+;base64,/, '');
  // start decoding
  const decodedFile = new Buffer(fileData, 'base64');
 	// extract extension
 	const fileExtension = encodedData.toString().slice(encodedData.indexOf('/') + 1, encodedData.indexOf(';'));
 	// root path
  var path = './public/uploads/' + got_arr.slice(-1)[0];
  // img path
  var img_path = path + '/' + String(got_arr[0]) + '.' + fileExtension;
  // make directory
  // directory not exists
  if(!fs.existsSync(path)) {
    // make directory
    fs.mkdir(path, function (err) {
      if(err){
        // making error
        console.log("e-901: label_directory_making_error");
      } else {
        // success
        console.log("904: label_directory_successfully_made");
        // generate img file
        fs.writeFile(img_path, decodedFile, (err) => {
          if(err){
            // making error
            console.log("e-903: label_image_making_error");
          } else {
            // success
            console.log("901: label_write_success");
            // lwip resizing
            sharp(img_path)
            .extract({ left: 0, top: 0, width: Math.round(Number(got_arr[1])), height: Math.round(Number(got_arr[2])) })
            .toFile(path + '/resized_' + String(got_arr[0]) + '.' + fileExtension, function(err) { 
              if(err){
                  // resize error
                  console.log("e-999: Resize err");
                } else {
                  // success
                  console.log("999: Resize success");
                  res.send("success");
                }
              }); 
            // init variable
            init();
          };
        });
      }
    });
  } else {
    console.log("905: label_directory_already_exists");
    // writing file
    fs.writeFile(img_path, decodedFile, (err) => {
      if(err){
        // making error
        console.log("e-903: label_image_making_error");
      } else {
        // success
        console.log("901: label_write_success");
        // lwip resize
        sharp(img_path)
        .extract({ left: 0, top: 0, width: Math.round(Number(got_arr[1])), height: Math.round(Number(got_arr[2])) })
        .toFile(path + '/resized_' + String(got_arr[0]) + '.' + fileExtension, function(err) { 
          　if(err){
              // resize error
              console.log("e-999: Resize err");
            　} else {
              // success
              console.log("999: Resize success");
              res.send("success");
            　}
          }); 
        // init variable
        init();
      };
    });
  }
});

// get payment data
app.post('/pay_response', function(req, res, next){
  // now time
  var now = getTime();
  // body
  var base = req.body;
  // insert payment result
  conn.query('INSERT INTO ?? (??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??, ??) VALUES ( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', ['payment_result', 'rand_id', 'shop_id', 'order_id', 'price', 'job_cd', 'access_id', 'access_pass', 'forwarded', 'approve', 'tranid', 'trandate', 'checkstring', 'newcardflg', 'paytype', 'cardno', 'created_at', 'updated_at', global_id, base.ShopID, base.OrderID, base.Amount, base.JobCd, base.AccessID, base.AccessPass, base.Forwarded, base.Approve, base.TranID, base.TranDate, base.CheckString, base.NewCardFlag, base.PayType, base.CardNo, now, now], function (err, results) {
    if(err){
      // INSERT error
      console.log(err);
    }
    // INSERT success
    if(results && base.Approve){
      console.log("success");
      // getID
      conn.query('SELECT * FROM ?? WHERE ?? = ?', ['payment', 'order_id', base.OrderID], function (err0, rows) {
        if(err0){
          // SELECT error
          console.log("e-035: DB Error -INSERT- -environment- : " + err);
        }
        if(rows){
        	conn.query('SELECT * FROM ?? WHERE ?? = ?', ['order', 'rand_id', rows[0].rand_id], function (err2, rows2) {
        		if(err2){
        			// SELECT error
              console.log("e-035: DB Error -INSERT- -environment- : " + err2);
            }
            if(rows2){
        			// mail body
              mailOptions.text = 'got order. please support it.\n\n[order]\n\ncustomer name：　' + rows2[0].username + '　\nLast name：　' + rows2[0].userruby + '\norganization：　' + rows2[0].usergroup + '\naddress：　' + rows2[0].useraddress + '\ntel：　' + rows2[0].usertel + '\nmail：　' + rows2[0].usermail + '\nto use date：　' + rows2[0].date + '\n\nprice：　' + base.Amount + 'yen（tax included）\n' + 'payment　：　done\n\n\nmanagement：　*****/login\n\nthank you';
		          // send mail
		          transporter.sendMail(mailOptions, function( error, info ){
                if(error){
		              // mail sending error
		              return console.log(error);
		            } else {
		              // success
		              console.log('Message sent: ' + info.response);
		              // update order table
		              conn.query('UPDATE ?? SET ?? = ? WHERE ?? = ?', ['order', 'paid', true, 'rand_id', global_id], function (err3, results3) {
                    if(err){
		                  // UPDATE error
		                  console.log("e-038: DB Error -UPDATE- -man_order- : " + err3);
		                } else {
		                  // UPDATE success
		                  console.log("224: man_order_update_success : " + results3);
		                  // redirect to thanks page
		                  res.redirect('*****/comp');
		                }
		              });
                }
              });
            }
          });
        }
      });
    } else {
      // payment error
      console.log("payment_error");
      // redirect
      res.redirect('*****');
    }  
  });
});

// report
app.post('/report', function(req, res, next){
  res.send(0);
});

// error handler
app.use(function (err, req, res, next) {
  // output logs
  console.error('err:', err);
  next(err);
});

// waiting port ↓--------------------------------------------------
app.listen(3000 , function () {
  console.log("902: port_open_success");
});


// general function ↓-----------------------------------------------------
// init
function init(){
	var global_id = "";
}

// get now time
function getTime(){
	// now time
	var date = new Date();
	date = date.toLocaleString().slice(0, 19).replace('T', ' ');
  var jikan = date.toLocaleString();
  return jikan;
}

// MD5 module
function md5hex(str) {
  const md5 = crypto.createHash('md5')
  return md5.update(str, 'binary').digest('hex')
}

// change date to ISO 8601 format
function replaceDate(d) {
  var year  = d.getFullYear();
  var month = d.getMonth() + 1;
  var day   = d.getDate();
  var hour  = ( d.getHours()   < 10 ) ? '0' + d.getHours()   : d.getHours();
  var min   = ( d.getMinutes() < 10 ) ? '0' + d.getMinutes() : d.getMinutes();
  var sec   = ( d.getSeconds() < 10 ) ? '0' + d.getSeconds() : d.getSeconds();
  return ( year + '-' + month + '-' + day + ' ' + hour + ':' + min + ':' + sec );
}

//　encrypt(AES192)
/*
function encrypt(planeText, passwrd){
  var cipher = crypto.createCipher('aes192', passwrd);
  cipher.update(planeText, 'utf8', 'hex');
  var cipheredText = cipher.final('hex');
  return cipheredText;
}

//　decrypt(AES192)
function decrypt(cipheredText, passwrd){
  var decipher = crypto.createDecipher('aes192', passwrd);
  decipher.update(cipheredText, 'hex', 'utf8');
  var dec = decipher.final('utf8');
  return dec;
}
*/