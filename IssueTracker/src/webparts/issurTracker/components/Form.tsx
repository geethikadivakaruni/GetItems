import * as React from 'react';
import { useState } from 'react';
//import ReactDOM from 'react-dom';
// import './App.css';
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField';
export function Form() {
	const [name , setName] = useState('');
	const [age , setAge] = useState('');
	const [email , setEmail] = useState('');
	const [password , setPassword] = useState('');
	const [confPassword , setConfPassword] = useState('');

	// function to update state of name with
	// value enter by user in form
	// const handleChange =(e)=>{setName(e.target.value);
	// }
	// // function to update state of age with value
	// // enter by user in form
	// const handleAgeChange =(e)=>{setAge(e.target.value);
	// }
	// // function to update state of email with value
	// // enter by user in form
	// const handleEmailChange =(e)=>{
	// setEmail(e.target.value);
	// }
	// // function to update state of password with
	// // value enter by user in form
	// const handlePasswordChange =(e)=>{
	// setPassword(e.target.value);
	// }
	// // function to update state of confirm password
	// // with value enter by user in form
	// const handleConfPasswordChange =(e)=>{
	// setConfPassword(e.target.value);
	// }
	// // below function will be called when user
	// // click on submit button .
	const handleSubmit=(e:any)=>{
		// if 'password' and 'confirm password'
		// does not match.
		alert("password Not Match");
	

	}
return (
	<div className="App">
	<header className="App-header">
	<form onSubmit={(e) => {handleSubmit(e)}}>
	{/*when user submit the form , handleSubmit()
		function will be called .*/}
	<h2> Geeks For Geeks </h2>
	<h3> Sign-up Form </h3>
    <TextField label="Standard" />
        <TextField label="Disabled" disabled defaultValue="I am disabled" />
        <TextField label="Read-only" readOnly defaultValue="I am read-only" />
        <TextField label="Required " required />
        <TextField ariaLabel="Required without visible label" required />
        <TextField label="With error message" errorMessage="Error message" />
		<input type="password" value={confPassword} required 
			 /><br/>
				{/* when user write in confirm password input box ,
					handleConfPasswordChange() function will be called.*/}
		<input type="submit" value="Submit"/> 
	</form>
	</header>
	</div>
);
}

export default Form;
