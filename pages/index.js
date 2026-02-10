// pages/index.js
import { useEffect, useState } from 'react';
import { Country, State, City } from 'country-state-city';

export default function Home() {
  // Only allow customers to review/edit these fields.
  const editableFields = [
    'Double Check No',
    'Full Name',
    'Phone Number',
    'Country/Region Code',
    'State/Province/Region',
    'City',
    'Address1',
    'Address2',
    'Zip Code',
    'Email'
  ];

  const [token, setToken] = useState('');
  const [address, setAddress] = useState(null);
  const [message, setMessage] = useState('');

  const selectedCountry = address?.['Country/Region Code'] || '';
  const selectedState = address?.['State/Province/Region'] || '';
  const countries = Country.getAllCountries();
  const states = selectedCountry ? State.getStatesOfCountry(selectedCountry) : [];
  const cities = selectedCountry && selectedState
    ? City.getCitiesOfState(selectedCountry, selectedState)
    : [];

useEffect(() => {
  try {
    const params = new URLSearchParams(window.location.search);
    const urlToken = params.get('token') || '';
    if(urlToken){
      setToken(urlToken);
      loadAddress(urlToken);
    }
  } catch (err) {
    console.error('Invalid URL token', err);
    setMessage('Invalid token in URL');
  }
}, []);


  const loadAddress = async (t) => {
    const currentToken = t || token;
    if(!currentToken) return;

    setMessage('');
    const res = await fetch(`/api/getAddress?token=${currentToken}`);
    const data = await res.json();
    if(res.ok){
      setAddress(data);
    } else {
      setMessage(data.error);
      setAddress(null);
    }
  }

  const submitAddress = async () => {
    setMessage('');
    const res = await fetch('/api/updateAddress', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ token, address })
    });
    const data = await res.json();
    if(res.ok){
      setMessage('Your information has been updated.');
    } else {
      setMessage(data.error);
    }
  }

  if(!address){
    return (
      <div style={{padding:20}}>
        {!token && <h2>No token found in URL</h2>}
        {token && <p>Loading your address...</p>}
        <p>{message}</p>
      </div>
    )
  }

  const renderField = (k) => {
    if (k === 'Country/Region Code') {
      return (
        <select
          value={selectedCountry}
          onChange={(e) =>
            setAddress({
              ...address,
              [k]: e.target.value,
              'State/Province/Region': '',
              City: ''
            })
          }
          style={{width: 320}}
        >
          <option value="">Select country</option>
          {countries.map((c) => (
            <option key={c.isoCode} value={c.isoCode}>
              {c.name}
            </option>
          ))}
        </select>
      );
    }

    if (k === 'State/Province/Region') {
      return (
        <select
          value={selectedState}
          onChange={(e) =>
            setAddress({
              ...address,
              [k]: e.target.value,
              City: ''
            })
          }
          style={{width: 320}}
        >
          <option value="">Select state/province</option>
          {states.map((s) => (
            <option key={s.isoCode} value={s.isoCode}>
              {s.name}
            </option>
          ))}
        </select>
      );
    }

    if (k === 'City') {
      if (cities.length === 0) {
        return (
          <input
            value={address[k] || ''}
            onChange={(e) => setAddress({ ...address, [k]: e.target.value })}
            placeholder="Enter city/town"
            style={{width: 300}}
          />
        );
      }

      return (
        <select
          value={address[k] || ''}
          onChange={(e) => setAddress({ ...address, [k]: e.target.value })}
          style={{width: 320}}
        >
          <option value="">Select city</option>
          {cities.map((c) => (
            <option key={c.name} value={c.name}>
              {c.name}
            </option>
          ))}
        </select>
      );
    }

    return (
      <input
        value={address[k] || ''}
        onChange={(e) => setAddress({ ...address, [k]: e.target.value })}
        style={{width: 300}}
      />
    );
  };

  return (
    <div style={{padding:20}}>
      <img
        src="/image.png"
        alt="Thank you"
        style={{maxWidth: 480, width: '100%', marginBottom: 16}}
      />
      <h2>Please Confirm or Edit your address</h2>
      {editableFields
        .filter((k) => k in address)
        .map((k) => (
        <div key={k}>
          <label>{k}: </label>
          {renderField(k)}
        </div>
      ))}
      <button onClick={submitAddress}>Submit</button>
      <p style={{marginTop: 8, color: '#555'}}>
        You can submit changes up to 3 times. Further submissions will be rejected.
      </p>
      <p>{message}</p>
    </div>
  )
}
