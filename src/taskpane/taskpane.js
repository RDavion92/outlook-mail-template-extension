/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

var email;

var elist;
var stemplate;


Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {

    console.log (info);
    
    email = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
    console.log (email);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("importtemplate").onclick = importtemplate;

    console.log(Office.context.mailbox.restUrl )

    elist = document.querySelector("#list")
    stemplate = document.querySelector("#template").outerHTML;
    elist.removeChild(document.querySelector("#template"));

    var data = { email: email }
    $.post('https://rdavion92.com/api/outlooktemplates', JSON.stringify(data), result => { listtemplate_callback(result) }, 'json' );
  }
});


function listtemplate_callback(result) 
{
  var _templates = result.templates
  console.log( _templates ) 

  elist.innerHTML = ""

  _templates.forEach( template => {

    if (template.to && template.to.length > 0) {
      template.to = template.to.split(',')
    } else {
      template.to = [] 
    }
    
    if (template.cc && template.cc.length > 0) {
      template.cc = template.cc.split(',')
    } else {
      template.cc = [] 
    }
    

    var ediv = document.createElement("div")
    ediv.innerHTML = stemplate;
    elist.appendChild(ediv);

    var b = ediv.querySelector(".btn-expand-template")
    b.innerHTML = template.subject;

    var b = ediv.querySelector(".div-template-body")
    b.innerHTML = template.body

    var b = ediv.querySelector(".btn-apply-template")
    b.onclick = function() { applyTemplate (template) } ;
    
    var b = ediv.querySelector(".btn-edit-template")
    b.onclick = function() { updateTemplate (template) } ;
    
    var b = ediv.querySelector(".btn-remove-template")
    b.onclick = function() { removeTemplate (template) } ;

  })
}




export async function applyTemplate( template ) {

  console.log(template)
  // Get a reference to the current message
  var item = Office.context.mailbox.item;

  item.to.setAsync(template.to); 
  item.cc.setAsync(template.cc); 

  item.subject.setAsync(template.subject)

  var shtml = '';
  if ( template.body.join ) 
  {
    shtml = template.body.join('');
  }
  else 
  {
    shtml = template.body;
  }

  item.body.prependAsync( shtml , {coercionType: Office.CoercionType.Html} )
}


export async function removeTemplate( template ) 
{
  var data = {
    action: 'remove',
    email: email,
    unid: template.unid
  }
  $.post('https://rdavion92.com/api/outlooktemplates', JSON.stringify(data), result => { importtemplate_callback(result) }, 'json' );
}

export async function updateTemplate(template) 
{
  template.action = 'import'
  getFormDatas(template, function(data) {

    $.post('https://rdavion92.com/api/outlooktemplates', JSON.stringify(data), result => { importtemplate_callback(result) }, 'json' );
        
  })
}



function uuidv4() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}


function importtemplate_callback(result) 
{
  console.log(result);
  var data = { email: email }
  $.post('https://rdavion92.com/api/outlooktemplates', JSON.stringify(data), result => { listtemplate_callback(result) }, 'json' );
}


function getFormDatas( data, callback ) {
  var item = Office.context.mailbox.item;
  data.to = []
  data.cc = []
  item.to.getAsync( to => {
    to.value.forEach( o => { data.to.push(o.emailAddress.toLowerCase() )})
    item.cc.getAsync( cc => {
      cc.value.forEach( o => { data.cc.push(o.emailAddress.toLowerCase())})
      item.subject.getAsync( subject => {
        data.subject = subject.value
        item.body.getAsync( Office.CoercionType.Html, body => {
          data.body = body.value
          data.to = data.to.join(',')
          data.cc = data.cc.join(',')
          console.log(data)
          callback(data);
        })
      })
    })
  })
}


export async function importtemplate() 
{
  var data = {
    action: 'import',
    email: email,
    unid: uuidv4(),
    to: [],
    cc: [],
    subject: ''
  }
  getFormDatas(data, function(data) {

    $.post('https://rdavion92.com/api/outlooktemplates', JSON.stringify(data), result => { importtemplate_callback(result) }, 'json' );
        
  })
}

