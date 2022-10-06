import * as React from 'react';
import {Component} from 'react'

export class Form extends Component{


render()
{
    return (
<div className='manage-app'>
<form id= "add-app">

<label>Title : </label>
<input type="text"> </input>

<label> Priority : </label>
<input type="text" ></input>

<label>Status </label>
<input ></input>
<label>Assigned to </label>
<input ></input>
<label>Datereported </label>
<input ></input>
<label>Images </label>
<input ></input>
<label>Issue source </label>
<input ></input>
<label>Amount </label>
<input ></input>
<label>Images </label>
<input ></input>


<button>Create</button>
</form>
        </div>
    );
}

}
export default Form 