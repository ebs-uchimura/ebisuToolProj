// ■ global variable
let globalRewrite
let globalCounter = 0
let globalFormCounter = 0
let globalImgCounter = 0
let globalImgArray = []

// ■ define template
// alchohol
const alchoholPage = {
  template: '#alchohol',
  methods: {
    push: function(n) {
      // badge
      this.$emit('enable-badge', 0)
      // page transfer
      this.$emit('sync-index', 0)
      // tmpDB
      this.$emit('insert-tmpdb', 1, n)
      // select page
      this.$emit('select-page', 0, n)
      // sort labels
      this.$emit('sort-label', 0)
    },
    tokusho: function() {
      this.$emit('tokusho', true)
    }
  }
}

// select product
const productPage = {
  template: '#product',
  props: ['page', 'prd_db'],
  methods: {
    push: function(n) {
      // badge
      this.$emit('enable-badge', 1)
      // page transfer
      this.$emit('sync-index', 1)
      // tmpDB
      this.$emit('insert-tmpdb', 2, n)
      // select page
      this.$emit('select-page', 1, n)
      // make labels
      this.$emit('make-labels')
    }
  }
}

// product detail
const detailPage = {
  template: '#detail',
  props: ['page', 'end_flg', 'prd_db'],
  methods: {
    push: function() {
      // badge
      this.$emit('enable-badge', 2)
      // page transfer
      this.$emit('sync-index', 2)
      // tmpDB
      this.$emit('insert-tmpdb', 3)
    }
  }
}

// prepare
const preparePage = {
  template: '#prepare',
  props: ['newEvent','selectedMethod', 'selectedBool1', 'selectedBool2'],
  methods: {
    push: function() {
      // badge
      this.$emit('enable-badge', 3)
      // page transfer
      this.$emit('sync-index', 3)
      // tmpDB
      this.$emit('insert-tmpdb', 4)
    },
    select: function(n) {
      // select method
      this.$emit('select-method', n)
    }
  },
  computed: {
  }
}

// select label
const labelPage = {
  template: '#label',
  props: ['labels', 'isExpanded'],
  methods: {
    push: function(n) {
      // tmpDB
      this.$emit('insert-tmpdb', 5, n)
      // badge
      this.$emit('enable-badge', 4)
      // page transfer
      this.$emit('sync-index', 4)
      // select label
      this.$emit('select-image', 4, n)
    },
    lb_sort: function(n) {
      // sort label
      this.$emit('sort-label', n)
    }
  }
}

// make label
const sendingPage = {
  template: '#sending',
  props: ['newEvent', 'form', 'page', 'dropImageUrls', 'ch', 'flg', 'flg2', 'sub', 'flgs'],
  methods: {
    push: function() {
      // localDB
      this.$emit('insert-db', 'font')
      this.$emit('insert-db', 'order')
      // badge
      this.$emit('enable-badge', 5)
      // page transfer
      this.$emit('sync-index', 5)
      // operate order
      this.$emit('add-cart', false)
      // upload image
      this.$emit('img_post', "work", '/upload')
      // clear style
      this.$emit('clear')
    },
    push2: function() {
      // localDB
      this.$emit('insert-db', 'font')
      this.$emit('insert-db', 'order')
      // badge
      this.$emit('disable-badge')
      // page transfer
      this.$emit('sync-index', -1)
      // operate order
      this.$emit('add-cart', false),
      // upload image
      this.$emit('img_post', "work", '/upload')
      // clear style
      this.$emit('clear')
    },
    push3: function() {
      // badge
      this.$emit('enable-badge', 3)
      // page transfer
      this.$emit('sync-index', 2)
    },
    font: function() {
      // set font main style
      this.$emit('font-input')
    },
    font2: function() {
      // set font sub style
      this.$emit('font-input2')
    },
    font3: function() {
      // set font other style
      this.$emit('font-input3')
    },
    size: function() {
      // set font main size
      this.$emit('font-size')
    },
    size2: function() {
      // set font sub size
      this.$emit('font-size2')
    },
    size3: function() {
      // set font other size
      this.$emit('font-size3')
    },
    vertical: function() {
      // set font vertical main size
      this.$emit('vertical')
    },
    vertical2: function() {
      // set font vertical sub size
      this.$emit('vertical2')
    },
    vertical3: function() {
      // set font vertical other size
      this.$emit('vertical3')
    },
    change_color: function(n, a) {
      // change font color
      this.$emit('ch_color', a, n)
    },
    tap: function(e) {
      // control click event
      e.preventDefault()
    },
    clear: function() {
      // clear font
      this.$emit('clear')
    }
  },
  computed: {
  }
}

// user registration page
const registPage = {
  template: '#regist',
  props: ['newEvent', 'regist', 'error', 'flg', 'form_no'],
  methods: {
    update: function() {
      // copy form
      this.$emit('send-brother')
      // send registred data
      this.$emit('update')
      // badge
      this.$emit('enable-badge', 6)
      // transfer page
      this.$emit('sync-index', 6)
      
    },
    push: function(n) {
      this.$emit('change-form', n)
    }
  },
  computed: {
  }
}

// confirm result
const resultPage = {
  template: '#result',
  props: ['newEvent', 'columns', 'table'],
  methods: {
    push: function(n) {
      if(n == 0) {
        this.$emit('sync-index', 5)
      } else {
        // badge
        this.$emit('enable-badge', 7)
        this.$emit('enable-badge', 8)
        // transfer page
        this.$emit('sync-index', 7)
        // edit cart info
        this.$emit('edit_cart')
        // localDB
        this.$emit('insert-db', 'user')
      }
    }
  }
}

// confirm cart information
const cartPage = {
  template: '#cart',
  props: ['tbl', 'gross', 'bt_flg', 'bt_flg2'],
  methods: {
    push: function() {
      // badge
      this.$emit('disable-badge')
      // return top page
      this.$emit('sync-index', -1)
    },
    go_order: function() {
      // transfer page
      this.$emit('sync-index', 5)
    },
    go_payment: function() {
      console.log("305: Order Comfirmation")
      // confirm order
      this.$emit('sync-index', 8)
    },
    delete_item: function(n) {
      // delete item
      this.$emit('delete-item', n)
    }
  }
}

// payment
const paymentPage = {
  template: '#payment',
  props: ['payment'],
  methods: {
    push: function() {
      // badge
      this.$emit('enable-badge', 8)
      // send payment information
      this.$emit('submit-payment', '/subpay')
    }
  }
}

// tokusho page
const tokushoPage = {
  template: '#tokusho',
  methods: {
    tokusho: function() {
      // page transfer
      this.$emit('tokusho', false)
    }
  }
}

// main
new Vue({
  el: '#app',
  template: '#main',
  // define data
  data() {
    return {
      // flg
      img_cre_flg: false,
      not_buy_flg: false,
      // label size
      lwidth: [],
      lheight: [],
      // window size
      wwidth: window.innerWidth,
      wheight: window.innerHeight,
      // cart display flg
      cart_on: true,
      // tmp localDB
      tmp_db: {
        title: '',
        sub_title: '',
        other_title: '',
        // img counter
        img_counter: 0,
        // selected alchohol ID
        selected_alc_id: 0,
        // selected item ID
        selected_prd_id: 0,
        // amount
        buyed_amt: 0,
        // selected lable ID
        selected_lb_id: 0,
        // labels
        labels: [],
        // auto flg
        lb_flg: false,
        // font information
        font: {
          main: {
            id: 0,
            color: 0
          },
          sub: {
            id: 0,
            color: 0
          },
          others: {
            id: 0,
            color: 0
          }
        }
      },
      // localDB
      local_db: {
        // order information
        order: [],
        // user information
        user: [],
        // font information
        font: [],
        // user environment information
        environment: []
      },
      // random key waiting flg
      create_flg: false,
      // POST send cargo
      cart_content: [],
      // order amount counter
      number: 0,
      // title bar display flg
      show: true,
      // tab bar display flg
      show_tab: true,
      // index of active page
      activeIndex: 0,
      // user ID
      rand_txt: "",
      // cart
      cart: {
        // title bar
        label2: 'shoppingcart',
        // tab bar
        label: 'cart',
        // page index
        index: 0,
        // template
        page: cartPage,
        // registration form display flg
        flg: false,
        props: {
          // cart table header
          columns: {
            id: 'ID',
            subject: 'subject',
            category: 'category',
            date: 'date'
          }
        },
        // distict key
        key: "cartPage",
        // visit flg
        visited: false
      },
      // main tab
      tabs: [
        // alchohol page
        {
          // tabbar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'select alchohol',
          // tab bar
          label: '①',
          // page index
          index: 1,
          // template
          page: alchoholPage,
          // registration form display flg
          flg: false,
          // distinct flg
          key: "alchoholPage",
          // visit flg
          visited: false,
          // URL
          path: '/alchohol'
        },
        // item page
        {
          // tab bar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'select type',
          // tab bar
          label: '②',
          // page index
          index: 2,
          // template
          page: productPage,
          // registration form display flg
          flg: false,
          props: {
            page: "",
            // item DB
            prd_db: {
              name: [], // name 
              price: [],　// price
              comment: [], // comment
              path: [] // image path
            }
          },
          // distinct key
          key: "productPage",
          // visit flg
          visited: false,
          // URL
          path: '/product'
        },
        // item detail page
        {
          // tab bar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'detail',
          // tab bar
          label: '③',
          // page index
          index: 3,
          // template
          page: detailPage,
          // registration form display flg
          flg: false,
          props: {
            page: "",
            end_flg: false,
            // itemDB
            prd_db: {
              name: [],　// name
              price: [],　// price
              comment: [],　// comment
              path: [], // path
              volume: [],　// volume
              alchohol: [],　// alcohol degree
              big_path: [], // big image path
              madein: [],　// made in
              long_comm: []　// long commnet
            }
          },
          // distinct key
          key: "detailPage",
          // visit flg
          visited: false,
          // URL
          path: '/detail'
        },
        // prepartion page
        {
          // tab bar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'enter label string',
          // tab bar
          label: '④',
          // page index
          index: 4,
          // template
          page: preparePage,
          // registration form display flg
          flg: false,
          // properties
          props: {
            selectedBool1: false,
            selectedBool2: true,
            selectedMethod: 1,
            newEvent: {　// titles
              title: '',
              sub_title: '',
              other_title: ''
            }
          },
          // distinct key
          key: "preparePage",
          // visit flg
          visited: false,
          // URL
          path: '/prepare'
        },
        // label page
        {
          // tab bar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'select base label',
          // tab bar 
          label: '⑤',
          // page index
          index: 5,
          // template
          page: labelPage,
          // color flg
          flg: false,
          // property
          props: {
            // labels
            labels: []
          },
          // distinct key
          key: "labelPage",
          // URL
          path: '/label'
        },
        // label making
        {
          // tabbar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: '',
          // tab bar 
          label: '⑥',
          // page index
          index: 6,
          // template
          page: sendingPage,
          // properties
          props: {
            flgs: {
              sub: false,
              other: false,
              all: false
            },
            // height of image
            img_height: 0,
            // color flg
            flg: 0,
            // color mode flg
            flg2: 0,
            //ac_active: true,
            sub: {
              // text flg
              sub_activeIndex: 0,
              sub_activeIndex2: 0
            },
            // text style
            form: {
              // font family
              items: [
              { text: '"Noto Serif JP", sans-serif', name: "1 Mincho", value: 1},
              { text: '"Noto Serif JP", sans-serif', name: "1 Micho", value: 2},
              { text: '"M PLUS Rounded 1c", sans-serif',name: "2 Rounded", value: 3},
              { text: '"Kosugi Maru", sans-serif', name: "3 Kosugi Maru", value: 4},
              { text: '"M PLUS 1p", sans-serif', name: "4 Gothic", value: 5}
              ],
              // max length
              m_l1: 1,
              m_l2: 1,
              m_l3: 1,
              // selected font
              selectedItem: 2,
              selectedItem2: 2,
              selectedItem3: 2,
              // last clicked
              lastClicked: 0,
              // texts store
              inputs: {
                text1: "",
                text2: "",
                text3: "" 
              },
              // label texts
              label_texts: {
                text1: "",
                text2: "",
                text3: "" 
              },
              // font size store
              volume: 60,
              volume2: 30,
              volume3: 30,
              // label canvas
              styleg: {
                width: '',
                height: '',
                backgroundImage: '',
                backgroundSize: ''
              },
              // style of main text
               style: {
                // position
                top: "0px",
                // font color
                color: "",
                // font family
                fontFamily: "",
                // font size
                fontSize: "60pt",
                // 
                writingMode: "horizontal-tb"
              },
              style2: {
                top: "100px",
                // font family
                fontFamily: "",
                // font size
                fontSize: "30pt",
                // Vertical writing
                writingMode: "horizontal-tb",
                // font color
                color: ""
              },
              style3: {
                top: "150px",
                // font family
                fontFamily: "",
                // font size
                fontSize: "30pt",
                // Vertical writing
                writingMode: "horizontal-tb",
                // font color
                color: ""
              },
              // switches
              switchOn: false,
              switchOn2: false,
              switchOn3: false
            },
            ch: {
              selectedEffect: [], // font effect
              here: "" // color mode           }
          },
          // distinct key
          key: "sendingPage",
          // visit flg
          visited: false,
          // URL
          path: '/sending'
        },
        // user registration
        {
          // tabbar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'customer',
          // tab bar
          label: '⑦',
          // page index
          index: 7,
          // template
          page: registPage,
          // registration form display flg
          flg: false,
          props: {
            form_no: 0,
            // form contents
            newEvent: {
              title: '',　// main title
              sub_title: '',　// sub title
              other_title: '', // other title
              meigi: '', // organizaiton name
              uname: '',　// name
              ruby: '',　// ruby
              zip: '', // zip address
              address: '', // address
              tel: '', // telephone
              mail: '', // mail address
              date: '',　// delivery date
              error: '', // error message
              check: false // checkbox
            },
          },
          // distinct key
          key: "registPage",
          // visit flg
          visited: false
        },
        // confirmation
        {
          // tabbar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'registered content',
          // tab bar
          label: '⑧',
          // page index
          index: 8,
          // template
          page: resultPage,
          // registration form display flg
          flg: false,
          // properties
          props: {
            // form contents
            newEvent: {
              title: '', // main title
              sub_title: '',　// sub title
              other_title: '', // other title
              meigi: '', // organization name
              uname: '', // name
              ruby: '', // ruby
              zip: '', // zip address
              address: '', // address
              tel: '',　// telephone
              mail: '',　// mail address
              date: '',　// delivery date
              error: '', // error message
              check: '' // checkbox
            },
          },
          // distict key
          key: "resultPage",
          // visit flg
          visited: false,
          // URL
          path: '/result'
        },
        // cart
        {
          // tabbar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'cart content',
          // tab bar
          label: '⑩',
          // page index
          index: 10,
          // template
          page: cartPage,
          // registration form display flg
          flg: false,
          // properties
          props: {
            // cart table
            tbl: [],
            // gross price
            gross: 0,
            // continue button display flg
            bt_flg: true,
            // goto order button display flg
            bt_flg2: false
          },
          // distinct key
          key: "cartPage",
          // visit flg
          visited: false,
          // URL
          path: '/cart'
        },
          // payment
        {
          // tabbar icon
          icon: 'fa-chevron-circle-right',
          // title bar
          label2: 'settlement',
          // tab bar
          label: '⑨',
          // page index
          index: 9,
          // template
          page: paymentPage,
          // registration form display flg
          flg: false,
          // properties
          props: {
            payment: {
              // shopID（fixed）
              ShopID: '****', 
              // orderID
              OrderID: '',
              // password（fixed）
              Password:  '****',
              // pass string
              ShopPassString: '',
              // result URL（fixed）
              RetURL: '****',
              // total price
              Amount: '',
              // creditcard（fixed）
              UseCredit: 1,
              // datetime
              DateTime: '',
              // distinction
              JobCd: '',
              // cancel URL
              CancelURL: '****'
            }
          },
          // distinct key
          key: "paymentPage",
          // visit kf
          visited: false,
          // URL
          path: '/payment'
        },
        // tokusho
        {
          label2: 'Specified Commercial Transaction Law',
          // tabbar
          label: '',
          // page index
          index: 11,
          // template
          page: tokushoPage
        }
      ]
    }
  }
},
// init
created() {
  globalImgArray.push(0)
  globalRewrite = this.rewrite
  history.pushState(0, null, null)
  // once
  if(!this.create_flg) {
    console.log("301: Vue created")
    // make random key
    var str = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    var len = 27
    var result = ""
    for(var i=0;i<len;i++) {
      result += str.charAt(Math.floor(Math.random() * str.length))
    }
    // random key
    this.rand_txt = result
    // post random key
    this.rand_post('/auth')
    this.create_flg = true
  }
},
// ignite after mount
mounted() {
  //window.addEventListener('scroll', this.handleScroll)
},
// page watcher
watch: {
},
// method
methods: {
  // rewrite
  rewrite(n) {
    this.activeIndex += n
    $("#datepicker").datepicker("hide");
  },

  // change form
  change_form: function(n) {
    this.tabs[6].props.newEvent.error = ""
    globalFormCounter += n
    console.log(this.tabs[6].props.newEvent.check)
    if((globalFormCounter == 2 && this.tabs[6].props.newEvent.uname == "") || (globalFormCounter == 6 && this.tabs[6].props.newEvent.tel == "") || (globalFormCounter == 10 && !this.tabs[6].props.newEvent.check)) {
      this.tabs[6].props.newEvent.error = "necessary item"
      globalFormCounter--
    } else {
      this.tabs[6].props.form_no += n
    }
  },

  // select mathod
  select_method: function(n) {
    // saver result
    this.tabs[3].props.selectedMethod = n
    if(n == 1) {
      this.tabs[3].props.selectedBool1 = false
      this.tabs[3].props.selectedBool2 = true
    } else {
      this.tabs[3].props.selectedBool1 = true
      this.tabs[3].props.selectedBool2 = false
    }
  },

  // tokusho
  tokusho: function(n) {
    if(n) {
      this.activeIndex = 10
    } else {
      this.activeIndex = 0
    }
  },

  // sort labels
  sort_label: function(n) {
    axios.post('/sort', {
      method : 'POST',　// post
      data   : { lb_no: n }　// body
    }).then(response => {
      // save got data
      this.tmp_db.labels = response.data
      console.log("303: random key post success")
      // goto label making page
      this.make_labels()
    })
  },

	// submit payment data
	subpay: function(url) {
    // route
		var base = this.tabs[9].props.payment
		axios.post(url, {
      method : 'POST',　// POST
      data   : { id: this.rand_txt, data: base } // ID and body
    }).then(response => {
      console.log("303: random key post success")
    })
	},

  // post random key
  rand_post: function(url) {
    console.log("302: random key post")
    // item array
    var item_arr = []
    axios.post(url, {
      method : 'POST',　// POST
      data   : { id: this.rand_txt } //ID
    }).then(response => {
      console.log("303: random key post success")
      // put into array
      response.data.forEach(function(value, key) {
        item_arr.push(value)
      })
      // put into productDB
      for(var i=0;i<item_arr.length;i++) {
        this.tabs[1].props.prd_db.name.push(item_arr[i].name)　// product name
        this.tabs[1].props.prd_db.price.push(item_arr[i].price)　// price
        this.tabs[1].props.prd_db.comment.push(item_arr[i].comment)　// comment
        this.tabs[1].props.prd_db.path.push(item_arr[i].path)　// image path
        this.tabs[2].props.prd_db.name.push(item_arr[i].name) // item name
        this.tabs[2].props.prd_db.price.push(item_arr[i].price.toLocaleString())　// price
        this.tabs[2].props.prd_db.path.push(item_arr[i].path) // image path
        this.tabs[2].props.prd_db.volume.push(item_arr[i].volume)　// volume
        this.tabs[2].props.prd_db.alchohol.push(item_arr[i].alchohol)　// alchohol degree
        this.tabs[2].props.prd_db.big_path.push(item_arr[i].big_path)　// big image path
        this.tabs[2].props.prd_db.madein.push(item_arr[i].madein) // madein
        this.tabs[2].props.prd_db.long_comm.push(item_arr[i].long_comm)　// long commnet
      }
    }).catch(err => {
      // error
      console.log('e-000: err:', err)
    })
  },

  // make label
  make_labels: function() {
    console.log("419: Editing label data")
    // chosen alchohol
    var alc_id = Number(this.tmp_db.selected_alc_id)
    // alhohol name
    var alc_names = ["shochu", "wine", "shochu"]
    // filename header
    var alc_head = alc_names[alc_id-1].slice(0, 1)
    // init array
    this.tabs[4].props.labels = []
    console.log("421: pass through flg")
    // enter label data
    for(var i = 0; i < this.tmp_db.labels.length; i++) {
      this.tabs[4].props.labels.push({
        url: "images/label/" + alc_names[alc_id-1] + "/" + alc_head + "_" + this.tmp_db.labels[i].img_url,　// image URL
        id: this.tmp_db.labels[i].id, // ID
        title: String(i+1) + ": " + this.tmp_db.labels[i].name, // label name
        attr: this.tmp_db.labels[i].attr,　// initial color
        explanation: this.tmp_db.labels[i].comment,　// comment
        ganre: this.tmp_db.labels[i].ganre　// genre
      })
    }
    console.log("422: data pushing finished")
  },

  // insert to tmpDB
  insert_tmpdb: function(a, n) {
    console.log("412: Insert into tmpdb ready")
    if(a == 1) {
      // selected alchohol ID
      this.tmp_db.selected_alc_id = n
    } else if(a == 2) {
      // selected item ID
      this.tmp_db.selected_prd_id = n
    } else if(a == 3) {
      // ordered amout
      this.tmp_db.buyed_amt = order_global_num
      order_global_num = 1
    } else if(a == 4) {
      this.tmp_db.title = this.tabs[3].props.newEvent.title
      this.tmp_db.sub_title = this.tabs[3].props.newEvent.sub_title
      this.tmp_db.other_title = this.tabs[3].props.newEvent.other_title
      this.tabs[5].props.form.inputs.text1 = this.tabs[3].props.newEvent.title
      this.tabs[5].props.form.inputs.text2 = this.tabs[3].props.newEvent.sub_title
      this.tabs[5].props.form.inputs.text3 = this.tabs[3].props.newEvent.other_title
      if(this.tmp_db.sub_title =="") {
        this.tabs[5].props.flgs.sub = true
      }
      if(this.tmp_db.other_title =="") {
        this.tabs[5].props.flgs.other = true
      }
      if(this.tmp_db.sub_title =="" && this.tmp_db.other_title =="") {
        this.tabs[5].props.flgs.all = true
      }
    } else if(a == 5) { 
      this.not_buy_flg = false
      // other
      this.tmp_db.lb_flg = false
      // label ID
      this.tmp_db.selected_lb_id = this.tmp_db.labels[n].id
      console.log("417: Insert into tmpdb finished")
    }
  },

  // insert into localDB
  insert_db: function(data) {
    // define variables
    var products = this.tabs[1].props.prd_db.name
    var prices = this.tabs[1].props.prd_db.price
    console.log("411: ready for inserting into localdb")
    // increment font ID
    this.tmp_db.font.main.id++
    this.tmp_db.font.sub.id++
    this.tmp_db.font.others.id++
    if(data == 'order' && !this.not_buy_flg) {
      // insert into localDB
      this.local_db.order.unshift({
        // ID
        id: this.rand_txt,
        // alchohol ID
        alchohol_id: this.tmp_db.selected_alc_id,
        // product ID
        product_id: this.tmp_db.selected_prd_id,
        // price
        price: prices[this.tmp_db.selected_prd_id-1],
        // item name
        name: products[this.tmp_db.selected_prd_id-1],
        // label ID
        label_id: this.tmp_db.selected_lb_id,
        // amount
        amount: this.tmp_db.buyed_amt,
        // auto flg
        lb_flg: this.tmp_db.lb_flg,
        // font ID
        font_id_main: this.tmp_db.font.main.id,
        font_id_sub: this.tmp_db.font.sub.id,
        font_id_others: this.tmp_db.font.others.id
      })
      console.log("423: inserting order data finished")
      this.not_buy_flg = false

      // font
    } else if(data == 'font') {
      this.local_db.font.unshift({
        // main text
        main: {
          // ID
          id: this.tmp_db.font.main.id,
          // content
          content: this.tabs[5].props.form.inputs.text1,
          // color
          color_id: this.tmp_db.font.main.color,
          // size
          size: this.tabs[5].props.form.volume,
          // font
          family: this.tabs[5].props.form.selectedItem,
          // x
          font_x: 0,
          //y
          font_y: 0
        },
        // sub text
        sub: {
          // ID
          id: this.tmp_db.font.sub.id,
          // content
          content: this.tabs[5].props.form.inputs.text2,
          // color
          color_id: this.tmp_db.font.sub.color,
          // size
          size: this.tabs[5].props.form.volume2,
          // font
          family: this.tabs[5].props.form.selectedItem2,
          // x
          font_x: 0,
          // y
          font_y: 0
        },
        // others
        others: {
          // ID
          id: this.tmp_db.font.others.id,
          // content
          content: this.tabs[5].props.form.inputs.text3,
          // color
          color_id: this.tmp_db.font.others.color,
          // size
          size: this.tabs[5].props.form.volume3,
          // font
          family: this.tabs[5].props.form.selectedItem3,
          // x
          font_x: 0,
          //y
          font_y: 0
        }
      })
      console.log("424: inserting font data finished")

    // environment
    } else if (data == "environment") {
      // user environment
      this.local_db.environment.unshift({
        // screen size
        screen_x: window.innerWidth,
        screen_y: window.innerHeight,
        // browser
        browser: _uac.browser,
        // device
        devise: _uac.device
      })
      console.log("425: inserting environment data finished")
    }
  },

  // edit cart
  edit_cart: function() {
    $("#imgs").html("")
    console.log("304: Write data to Cart")
    // hide order button
    if(this.number == 0) {
      console.log("409: Cart is empty")
    }
    for(var i=this.number-1;i>=0;i--) {
      $("#imgs").append('<img width="50%" id="thumb_' + String(globalImgArray[i]) + '" src="./uploads/' + this.rand_txt + '/resized_' + String(globalImgArray[i]) + '.png" />')
    }
      // item name
      var products = this.tabs[1].props.prd_db.name
      // price
      var prices = this.tabs[1].props.prd_db.price
      // item ID
      var p_id = 0
      var total_arr = []
      for(var i=0;i<this.local_db.order.length;i++) {
        // ID
        p_id = this.local_db.order[i].product_id
        // make empty obj
        this.tabs[8].props.tbl[i] = {
          name: "",
          labelid: 0,
          content: "",
          itemcount: 0,
          price: 0
        }
        // extract data from localDB
        this.tabs[8].props.tbl[i].name = products[p_id-1]
        this.tabs[8].props.tbl[i].labelid = this.local_db.order[i].label_id - 1
        // strings
        this.tabs[8].props.tbl[i].content = this.local_db.font[i].main.content
        // ordered amount
        this.tabs[8].props.tbl[i].itemcount = this.local_db.order[i].amount
        // price
        this.tabs[8].props.tbl[i].price = prices[p_id-1]
        // sum
        total_arr[i] = prices[p_id-1] * this.local_db.order[i].amount
      }
      // total price
      if(this.number > 0) {
        this.tabs[8].props.gross = sum(total_arr)
      }
      console.log("426: filling receipt table finished")
      if(this.activeIndex !== 9) {
        this.activeIndex = 8
      }
  },

  // copy form data
  send_brother: function() {
    console.log("408: Copy Form Data")
    // copy form data
    this.tabs[7].props.newEvent = this.tabs[6].props.newEvent
    this.tabs[7].props.newEvent.date = $("#datepicker").val()
  },

  // page transfer
  select: function(a, n) {
    console.log("321: Page forward")
    // page transfer
    this.tabs[a+1].props.page = n
  },

  // tab control
  getTab: function(e) {
    e.preventDefault()
    // init slider
    console.log("307: Tab action")
    // disable double click
    if(this.activeIndex < e.tabItem.index && !this.tabs[e.tabItem.index].visited) {
      console.log("308: Tab click cancel")
      // cancel
      e.cancel()
    }
    if(this.activeIndex == 8 && this.tabs[e.tabItem.index].visited && (e.tabItem.index == 4 || e.tabItem.index == 5 )) {
      console.log("308: Tab click cancel")
      // cancel
      e.cancel()
    }
    if(this.activeIndex > 8 && e.tabItem.index < 9　&& this.activeIndex !== 10) {
      console.log("308: Can't back after ordering")
      alert("You can't go back after settlement confirmed. If retry shopping, reload page.")
      // cancel
      e.cancel()
    }
    if(this.activeIndex == 10 && (e.tabItem.index == 6 || e.tabItem.index == 7 || e.tabItem.index == 8 || e.tabItem.index == 9)) {
      console.log("308: Tab click cancel")
      // cancel
      e.cancel()
    }
    if(this.activeIndex > 5 && e.tabItem.index == 5) {
      console.log("308: Tab click cancel")
      // cancel
      e.cancel()
    }
    if(this.activeIndex == 8 && e.tabItem.index !== 8) {
      console.log("308: Tab click cancel")
      // cancel
      e.cancel()
    } else if(this.activeIndex > e.tabItem.index && this.activeIndex > 5) {
      console.log("309: Back to shopping confirmation")
      // show dialog
      var res = confirm("continue shopipng？")
      // yes
      if(res) {
        console.log("310: Back to shopping")
        // delete all badges
        this.tabs.forEach(function(item) {
          item.badge = ""
        })
        // init event
        e.index = 0
        // init visit flg
        for(var i=0;i<this.tabs.length;i++) {
          this.tabs[i].visited = false
        }
        console.log("427: visited initializing finished")
        // no
      } else {
        // cancel
        e.cancel()
      }
    }
  },

  // image upload（POST)
  img_post: function(im, url) {
    this.img_cre_flg = false
    // increment
    this.tmp_db.img_counter++
    var w = window.innerWidth
    var h = window.innerHeight
    // wine
    if(this.tmp_db.selected_alc_id == 2) {
      // optimize size
      var p_size = calc_size(188, 254, false)
      console.log("890: img is wine")
      // optimize position
       if(w < 321) {
        // iphone5
        var sx = p_size[0] * 0.12
        var sy = p_size[1] * 0.05
      } else {
        var sx = p_size[0] * 0.1
        var sy = p_size[1] * 0.06
      }
    // shochu
    } else {
      // optimize size
      var p_size = calc_size(300, 254, true)
      console.log("891: img is shochu")
      if(w < 321) {
        //　iphone5
        var sx = p_size[0] * 0.07
        var sy = p_size[1] * 0.05
      } else {
        var sx = p_size[0] * 0.05
        var sy = p_size[1] * 0.06
      }
    }
    // header
    var x_header = [this.tmp_db.img_counter, p_size[0], p_size[1], this.rand_txt]
    console.log("316: Image Uploading")
    // label upload header
    var instance = axios.create({
      'responseType': 'arraybuffer',
      'headers': {
      'ContentType': 'image/png',
      'X-Custom-Header': x_header
      }
    })
    // optimize string position
    if(this.tabs[5].props.form.style.writingMode == "vertical-rl") {
      $("#text_p1").css('position', 'relative')
      $("#text_p1").css('left', sx + 'px')
    } else {
      $("#text_p1").css('position', 'relative')
      $("#text_p1").css('bottom', sy + 'px')
    }
    if(this.tabs[5].props.form.style2.writingMode == "vertical-rl") {
      $("#text_p2").css('position', 'relative')
      $("#text_p2").css('left', sx + 'px')
    } else {
      $("#text_p2").css('position', 'relative')
      $("#text_p2").css('bottom', sy + 'px')
    }
    if(this.tabs[5].props.form.style3.writingMode == "vertical-rl") {
      $("#text_p3").css('position', 'relative')
      $("#text_p3").css('left', sx + 'px')
    } else {
      $("#text_p3").css('position', 'relative')
      $("#text_p3").css('bottom', sy + 'px')
    }
    // html2canvas
    html2canvas(document.querySelector("#labelee_canvas"), {
      allowTaint: true,
      foreignObjectRendering: true,
      width: w,
      height: h,
      scrollX: 0,
      scrollY: 0,
      x: 0,
      y:0
    }).then(function(canvas) {
      console.log("316: Image Creating By html2canvas")
      // encode to BASE64
      var dat = base64Encode(canvas.toDataURL('image/png'))
      　// upload
      instance.post(url, { data : { sent_item : dat }
        }).then(res => {
          // success
          console.log("318: Image Uploading Success " + res)
          this.img_cre_flg = true
        }).catch(err => {
          // fail
          console.log("e-305: image upload failed " + err) 
      })
    })
  },

  // main horizen-vertical change
  vert: function() {
    console.log("322: Vertical change -main-")
    // when switch on
    if(!this.tabs[5].props.form.switchOn){
      // vertical
      this.tabs[5].props.form.style.writingMode = "vertical-rl"
    } else {
      // horizontal
      this.tabs[5].props.form.style.writingMode = "horizontal-tb"
    }
  },

  // sub horizen-vertical change
  vert2: function() {
    console.log("323: Vertical change -sub-")
    // when switch on
    if(!this.tabs[5].props.form.switchOn2){
      // vertical
      this.tabs[5].props.form.style2.writingMode = "vertical-rl"
    } else {
      // horizontal
      this.tabs[5].props.form.style2.writingMode = "horizontal-tb"
    }
  },

  // other horizen-vertical change
  vert3: function() {
    console.log("324: Vertical change -others-")
    // when switch on
    if(!this.tabs[5].props.form.switchOn3){
      // vertical
      this.tabs[5].props.form.style3.writingMode = "vertical-rl"
    } else {
      // horizontal
      this.tabs[5].props.form.style3.writingMode = "horizontal-tb"
    }
  },

  // text color
  ch_color: function(no, a) {
    console.log("325: Font Color changing")
    // get DOM
    var el1 = document.getElementById('text_p1')
    var el2 = document.getElementById('text_p2')
    var el3 = document.getElementById('text_p3')
    // initialize style
    el1.removeAttribute()
    el2.removeAttribute()
    el3.removeAttribute()
    // define monochrome mode
    var num = Number(this.tabs[5].props.ch.here)
    if(num == 1) {
      // white mode
      flg2 = 1
    } else {
      flg2 = 2
      console.log("428: color mode is white")
    }
    console.log("326: Font Color preparation")
    // insert into tmpDB
    this.tmp_db.font.main.color = no
    switch(no) {
      case 1:
        console.log("327: Font Color -black-")
        // black
        eval("el" + a + ".setAttribute('style', 'color: #000000;')")
        break
      case 2:
        console.log("328: Font Color -white-")
        // white
        eval("el" + a + ".setAttribute('style', 'color: #FFFFFF;')")
        break
      case 3:
        console.log("329: Font Color -red-")
        // red
        eval("el" + a + ".setAttribute('style', 'color: #FC2347;')")
        break
      case 4:
        console.log("330: Font Color -blue-")
        // blue
        eval("el" + a + ".setAttribute('style', 'color: #0000cd;')")
        break
      case 5:
        console.log("331: Font Color -green-")
        // green
        eval("el" + a + ".setAttribute('style', 'color: #228b22;')")
        break
      case 6:
        console.log("332: Font Color -pink-")
        // pink
        eval("el" + a + ".setAttribute('style', 'color: #ff1493;')")
        break
      case 7:
        console.log("333: Font Color -yellow-")
        // yellow
        eval("el" + a + ".setAttribute('style', 'color: #ffd700;')")
        break
      case 8:
        console.log("334: Font Color -orange-")
        // orange
        eval("el" + a + ".setAttribute('style', 'color: #ff8c00;')")
        break
      case 9:
        console.log("336: Font Color -violet-")
        // purple
        eval("el" + a + ".setAttribute('style', 'color: #A757A8;')")
        break
      case 10:
        console.log("337: Font Color -skyblue-")
        // sky blue
        eval("el" + a + ".setAttribute('style', 'color: #87ceeb;')")
        break
      case 11:
        console.log("339: Font Color -lightgreen-")
        // light green
        eval("el" + a + ".setAttribute('style', 'color: #90ee90;')")
        break
      case 12:
        console.log("340: Font Color -lightpink-")
        // light pink
        eval("el" + a + ".setAttribute('style', 'color: #ffb6c1;')")
        break
    }
  },

  // clear text
  clear: function() {
    console.log("377: Text Cleaning")
    // label no
    var num = Number(this.tabs[5].props.ch.here)
    // get text DOM
    var el = document.getElementById('text_p1')
    var el2 = document.getElementById('text_p2')
    var el3 = document.getElementById('text_p3')
    // clear font style
    el.removeAttribute('style')
    el2.removeAttribute('style')
    el3.removeAttribute('style')
    // remove attributes
    this.tabs[5].props.form.style.writingMode = ""
    this.tabs[5].props.form.style2.writingMode = ""
    this.tabs[5].props.form.style3.writingMode = ""
    if(num == 1){
      // clear white
      console.log("378: Text Cleaning -white-")
      el.setAttribute('style', 'color: #FFF;')
      el2.setAttribute('style', 'color: #FFF;')
      el3.setAttribute('style', 'color: #FFF;')
      this.tabs[5].props.form.selectedColor = 2
      this.tabs[5].props.form.selectedColor2 = 2
      this.tabs[5].props.form.selectedColor3 = 2
    } else {
      // clear black
      console.log("379: Text Cleaning -black-")
      el.setAttribute('style', 'color: #000;')
      el2.setAttribute('style', 'color: #000;')
      el3.setAttribute('style', 'color: #000;')
      this.tabs[5].props.form.selectedColor = 1
      this.tabs[5].props.form.selectedColor2 = 1
      this.tabs[5].props.form.selectedColor3 = 1
    }
    // remove text
    this.tabs[5].props.form.inputs.text1 = ""
    this.tabs[5].props.form.inputs.text2 = ""
    this.tabs[5].props.form.inputs.text3 = ""
    // initialize font size
    this.tabs[5].props.form.volume = 60
    this.tabs[5].props.form.volume2 = 30
    this.tabs[5].props.form.volume3 = 30
    // initialize font effect
    this.tabs[5].props.ch.selectedEffect = []
    // initialize slider
    this.tabs[5].props.form.switchOn = false
    this.tabs[5].props.form.switchOn2 = false
    this.tabs[5].props.form.switchOn3 = false
  },

  // main font style
  getfont: function() {
    console.log("380: Font styling -main-")
    var i = Number(this.tabs[5].props.form.selectedItem) - 1
    var fontname = this.tabs[5].props.form.items[i].text
    this.tabs[5].props.form.style.fontFamily = fontname
  },

  // sub font style
  getfont2: function() {
    console.log("381: Font styling -sub-")
    var i = Number(this.tabs[5].props.form.selectedItem2) - 1
    var fontname = this.tabs[5].props.form.items[i].text
    this.tabs[5].props.form.style2.fontFamily = fontname
  },

  // other font style
  getfont3: function() {
    console.log("382: Font styling -others-")
    var i = Number(this.tabs[5].props.form.selectedItem3) - 1
    var fontname = this.tabs[5].props.form.items[i].text
    this.tabs[5].props.form.style3.fontFamily = fontname
  },

  // change font size (main)
  sizechange: function() {
    console.log("383: Font size change -main-")
    var vl = this.tabs[5].props.form.volume
    this.tabs[5].props.form.style.fontSize = vl + "pt"
    this.tabs[5].props.form.m_l1 = window.innerWidth / vl
  },

  // change font size (sub)
  sizechange2: function() {
    console.log("384: Font size change -sub-")
    var vl2 = this.tabs[5].props.form.volume2
    this.tabs[5].props.form.style2.fontSize = vl2 + "pt"
    this.tabs[5].props.form.m_l2 = window.innerWidth / vl2
  },

  // change font size (other)
  sizechange3: function() {
    console.log("385: Font size change -others-")
    var vl3 = this.tabs[5].props.form.volume3
    this.tabs[5].props.form.style3.fontSize = vl3 + "pt"
    this.tabs[5].props.form.m_l3 = window.innerWidth / vl3
  },

  // add budge
  enableit : function(page) {
    console.log("386: Tab badge adding")
    this.tabs[page].badge = "済"
    this.tabs[page].visited = true
  },

  // clear budge
  disableall: function() {
    console.log("387: Tab badge clearing all")
    this.tabs.forEach(function(item) {
      item.badge = ""
      item.visited = false
    })
  },

  // synchronize page
  syncindex : function(no) {
    globalCounter++
    history.pushState(globalCounter, null, null)
    this.tabs[2].props.end_flg = true
    console.log("388: Page synchronizing")
    if(this.activeIndex == 5) {
      // hide title bar
      this.show = false
    }
    if(no == 2) {
      // hide cart icon
      this.cart_on = true
      // show tabbar
      this.show_tab = true
      // global
      if(order_global_num > 0 && typeof(order_global_num) == "number") {
        // transfer page
        this.activeIndex = 3
      } else {
        alert("enter over 1");
      }
    } else if(no == 3) {
      if(!this.tabs[3].props.newEvent.title) {
        alert("enter main character.")
      } else {
        this.activeIndex = 4
      }
    } else if(no == 4) {
      order_global_num = 1
      this.cart_on = true
      // hide title bar
      this.show = false
      // hide tabbar
      this.show_tab = false
      // hide edge
      this.tabs[5].props.flg = false
      // transfer page
      if(this.tabs[3].props.selectedMethod == 1) {
        this.show = true
        this.cart_on = true
        // show tabbar
        this.show_tab = true
        this.activeIndex = 6
        this.addcart(false)
        // upload image 
        this.img_post("work", '/upload')
        this.insert_db("font")
        this.insert_db("order")
      } else {
        this.activeIndex = 5
      }
      console.log("344: Page sync to label edit to " + this.activeIndex)
    } else if(no == 5) {
      this.tabs[5].props.isButton = false
      // show title bar
      this.show = true
      this.cart_on = true
      // show tabbar
      this.show_tab = true
      // transfer page
      this.activeIndex = 6
      console.log("389: Page sync to regist to " + this.activeIndex)
    } else if(no == 6) {
      // show tabbar
      this.show_tab = true
      // transfer page
      this.activeIndex = 7
    } else if(no == 7) {
      // show title bar
      this.show = true
      this.cart_on = true
      // show tabbar
      this.show_tab = true
      // transfer page
      this.activeIndex = 8
      this.tabs[8].props.bt_flg = false
      console.log("399: Page sync to cart to " + this.activeIndex)
    } else if(no == 8) {
      if(this.number > 0) {
        // show title bar
        this.show = true
        // show tabbar
        this.show_tab = true
        this.cart_on = true
        // transfer page
        this.activeIndex = 9
        // POST data
        this.con_post('/post_json')
        // goto payment page
        this.gopay()
      } else {
        var res = confirm("cart is empty, go back?")
        if(res) {
          // goto top
          this.activeIndex = 0
        } else {
          return false
        }
      }
    } else {
      // show title bar
      this.show = true
      this.cart_on = true
      // show tabbar
      this.show_tab = true
      // transfer page
      this.activeIndex = no + 1
      console.log("391: Page sync to others to " + this.activeIndex)
    }
  },

  // POST
  con_post: function(url) {
    // form info
    var user_itm = this.tabs[7].props.newEvent
    // order info
    var order_itm = this.local_db.order
    // font info
    var font_itm = this.local_db.font
    // environment info
    var environment_itm = this.local_db.environment
    // make JSON
    var itm = {user: user_itm, order: order_itm, font: font_itm, environment: environment_itm, total: this.tabs[8].props.gross}
    // submit
    axios.post(url, {
      method : 'POST',
      data   : { sent_item : itm, uid : this.rand_txt}
    }).then(response => console.log("888: Status: " + response.status));
  },

  // set label image
  selimage: function(a, n) {
    console.log("392: Font initializing")
    this.tabs[5].props.form.label_texts.text1 = this.tabs[3].props.newEvent.title
    if(this.tabs[3].props.newEvent.sub_title == "") {
      this.tabs[5].props.form.label_texts.text2 = "sub character"
    } else {
      this.tabs[5].props.form.label_texts.text2 = this.tabs[3].props.newEvent.sub_title
    }
    if(this.tabs[3].props.newEvent.other_title == "") {
      this.tabs[5].props.form.label_texts.text3 = "other character"
    } else {
      this.tabs[5].props.form.label_texts.text3 = this.tabs[3].props.newEvent.other_title
    }

    // image size
    var i_flg = true
    var size_x = window.innerWidth
    var size_y = window.innerHeight
    // wine
    if(this.tmp_db.selected_alc_id == 2) {
      var p_size = calc_size(188, 254, false)
      i_flg = false
      this.lwidth.unshift(p_size[0])
      this.lheight.unshift(p_size[1])
      console.log("890: img is wine")
    // shochu
    } else {
      var p_size = calc_size(300, 254, true)
      i_flg = true
      this.lwidth.unshift(size_x)
      this.lheight.unshift(p_size[1])
      console.log("891: img is shochu")
    }
    // set canvas style
    if(i_flg) {
      $("#labelee_canvas").width(size_x)　// width
      $("#labelee_canvas").height(p_size[1])　// height
      $("#labelee_canvas").css("background-size", size_x + "px" + " " + p_size[1] + "px")　// bg image size
      $("#labelee_canvas").css("left", "0px")  // position
    } else {
      $("#labelee_canvas").width(p_size[0])　// width
      $("#labelee_canvas").height(p_size[1])　// height
      $("#labelee_canvas").css("background-size", p_size[0] + "px" + " " + p_size[1] + "px")　// bg image size
      $("#labelee_canvas").css("position", "relative")　// relative position
      $("#labelee_canvas").css("left", (size_x - p_size[0])/2 + "px")　// position
    }
    // detect color 
    this.tabs[5].props.ch.here = String(this.tabs[4].props.labels[n].attr)
    // whiteblack
    var num = Number(this.tabs[5].props.ch.here)
    // get text DOM
    var el = document.getElementById('text_p1')
    var el2 = document.getElementById('text_p2')
    var el3 = document.getElementById('text_p3')
    if(num == 1) {
      console.log("394: Font is white")
      // set white
      el.style.color = "#FFF"
      el2.style.color = "#FFF"
      el3.style.color = "#FFF"
    } else if(num == 3 || num == 4) {
      console.log("395: Font is mixture")
      // set white and black
      el.style.color = "#FFF"
      el2.style.color = "#000"
      el3.style.color = "#000"
    } else if (num == 2) {
      console.log("396: Font is black")
      // set black
      el.style.color = "#000"
      el2.style.color = "#000"
      el3.style.color = "#000"
    }
    // transfer page
    this.tabs[a+1].props.page = n
    var base_url = '****'
    // set URL
    if(this.backup_img_urls) {
      this.tabs[5].props.form.styleg.backgroundImage = "url(" + base_url + "/upload/" + $("#r_id").val() + "/" + $("#r_num").val() + ".png)"
    } else {
      this.tabs[5].props.form.styleg.backgroundImage = "url(" + this.tabs[4].props.labels[n].url + ")"
    }
  },

  // registration button
  update () {
    console.log("398: Go to Registration Form")
    // get DOM
    const form = document.getElementById('regist_form')
    const riyu = document.getElementById('riyu')
    // form valudation
    if (!form.checkValidity()) {
      console.log("e-302: Validation error")
      // form message
      this.tabs[6].props.newEvent.error = 'enter all'
    } else {
      console.log("400: Go to Registration Form")
      // transfer page
      this.activeIndex = 7
      // add badge
      this.tabs[6].badge = "done"
    }
  },

  // add cart
  addcart: function(flg) {
    globalImgCounter++
    globalImgArray.unshift(globalImgCounter)
    console.log("404: Cart Increment")
    // item increment
    this.number++
    // show flg
    this.tabs[8].props.bt_flg2 = true
    // init number
    $(".spinner").val(1)
    if(flg) {
      this.tabs[10].props.boxes++
      if(this.tabs[10].props.boxes > 3) {
        alert("no box")
        this.tabs[10].props.boxes = 3
        this.activeIndex = 3
        this.number--
        this.not_buy_flg = true
      }
    } else {
      this.not_buy_flg = false
    }
  },

  // order information
  gopay: function() {
    console.log("345: Go to GMO payment site")
    // data
    var base = this.tabs[9].props.payment
    // ★ goto GMO payment
    // order ID
    base.OrderID = this.rand_txt
    // total price
    base.Amount = this.tabs[8].props.gross
    // date
    base.DateTime = getTimeforpayment()
    // distinction
    base.JobCd = 'CHECK'
    // identification str
    base.ShopPassString = md5(base.ShopID + "|" + base.OrderID + "|" + base.Amount + "||" + base.Password + "|" + base.DateTime)
  },

  // delete cart item
  delete_item: function(index) {
    console.log("346: cart item deleted")
    if (this.boxes > 0 && this.tabs[8].props.tbl[index].labelid == "trust") {
      this.tabs[10].props.boxes--
      this.not_buy_flg = false
    }
    // calc gross
    this.tabs[8].props.gross -= this.tabs[8].props.tbl[index].price * this.tabs[8].props.tbl[index].itemcount
    // delete from local DB
    this.local_db.order.splice(index, 1)
    // delete from table
    this.tabs[8].props.tbl.splice(index, 1)
    // decrement item
    this.number--
    globalImgArray.slice(index, 1)
  }
},

// compute
computed: {
}
})

// init num
var order_global_num = 1

window.addEventListener('DOMContentLoaded', function() {
  $(window).on('popstate', function (e) {
    if(globalCounter > e.originalEvent.state){
      globalRewrite(-1)
    } else {
      globalRewrite(1)
    }
    globalCounter = e.originalEvent.state
  })
})

$(function() {
  // auto zip code
  AjaxZip3.onSuccess = function() {
    AjaxZip3.zip2addr($("#zip"),'','addr11','addr11');
  }
  console.log("405: Jquery started")
  $("#sending").on('touchmove.noScroll', function(e) {
    e.preventDefault();
  });
  $(".text1").on('touchmove.noScroll', function(e) {
    e.preventDefault();
  });

  //　calender
  $.datepicker.setDefaults( $.datepicker.regional[ "ja" ] );
  var week = nextweek();
  var week_slash = getNowYMD(week);
  $("#datepicker").val(week_slash);
  $("#datepicker").datepicker({
    defaultDate: week,
    minDate: week,
    autoSize: true
  }).on("dp.change", function(e) {
    e.date.format(e.date._f);
  });
  console.log("420: calender started");
  // prohibit resize（iphone用）
  var windowWidth = $(window).width();

  // resizing
  $(window).resize(function() {
    var ww = $(window).width();
    if(windowWidth !== ww) {
      console.log("415: Avoid Resizing Completed");
      windowWidth = ww;
    }
  });

  $(".spinner").on('input', function() {
    // hankanku exchange
    var halfVal = $(this).val().replace(/[！-～]/g, function (tmpStr) {
        // shift character code
        return String.fromCharCode(tmpStr.charCodeAt(0) - 0xFEE0);
      }
    );
    // delete unessesary code
    $(this).val(halfVal.replace(/[^0-9]/g, ''));
  });

  // get amount
  $(".spinner").spinner({
    max: 99,
    min: 1,
    step: 1
  });

  $(".spinner").spinner({
    change: function(event, ui) {
      order_global_num = $(this).spinner("value");
    }
  });

  // submit form
  $("#payment_part").click(function() {
    $("#pay_button").submit();
  });
});

// general function ↓ -----------------------------------------------------------------
// BASE64 encode
function base64Encode(imgData) {
  // encode imgdata to base64
  const base64Encoded = imgData.toString('base64');
  return base64Encoded;
}

// generate zero
function zeroPadding(NUM, LEN) {
  return ( Array(LEN).join('0') + NUM ).slice( -LEN );
}

// calc total
var sum = function(arr) {
  return arr.reduce(function(prev, current, i, arr) {
    return prev+current;
  });
};

// nowtime
function getTime(){
  var now = new Date();
  var year = now.getFullYear();
  var mon = now.getMonth() + 1;
  var day = now.getDate();
  var date = year + "/" + mon + "/" + day;
  var jikan = date.toLocaleString();
  return jikan;
}

// get now time for payment
function getTimeforpayment(){
  var now = new Date();
  var year = String(now.getFullYear());
  var mon = now.getMonth()+1
  if (mon < 10) {
    mon = '0' + String(mon);
  }
  var day = now.getDay();
  if (day < 10) {
    day = '0' + String(day);
  }
  var hour = now.getHours();
  if (hour < 10) {
    hour = '0' + String(hour);
  }
  var min = now.getMinutes();
  if (min < 10) {
    min = '0' + String(min);
  }
  var sec = String('0' + now.getSeconds()).slice(-2);
  var date = year + mon + day + hour + min + sec;
  return date;
}

// calculate canvas size
function calc_size(px, py, flg) {
  // defalut height
  var base_y = 270;
  // set size
  var size_x = window.innerWidth;
  var size_y = window.innerHeight;
  // after width
  if(flg) {
    // calc variables
    var calc = size_x / px;
    var dec_x = size_x;
    var dec_y= py * calc;
    def_y = dec_y;
    // return result
    return [dec_x, dec_y];

  } else {
    // calc variables
    var calc2 = base_y / py;
    var des_x = px * calc2;
    var des_y = base_y;
    // return result
    return [des_x, des_y];
  }
}

// next week
function nextweek() {
    var today = new Date();
    var nextweek = new Date(today.getFullYear(), today.getMonth(), today.getDate()+7);
    return nextweek;
}

// get YMD
function getNowYMD(dt) {
  var y = dt.getFullYear();
  var m = ("00" + (dt.getMonth()+1)).slice(-2);
  var d = ("00" + dt.getDate()).slice(-2);
  var result = y + "/" + m + "/" + d;
  return result;
}