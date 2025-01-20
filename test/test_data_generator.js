import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardHeader, CardTitle, CardContent } from '@/components/ui/card';

const TestDataGenerator = () => {
  const [output, setOutput] = useState('');
  
  // German street names
  const streets = [
    'Hauptstraße', 'Schulstraße', 'Bahnhofstraße', 'Gartenstraße', 'Kirchstraße',
    'Waldstraße', 'Dorfstraße', 'Bergstraße', 'Lindenstraße', 'Mozartstraße'
  ];
  
  // German cities
  const cities = [
    'Berlin', 'Hamburg', 'München', 'Köln', 'Frankfurt',
    'Stuttgart', 'Düsseldorf', 'Dresden', 'Leipzig', 'Hannover'
  ];
  
  // German last names
  const lastNames = [
    'Müller', 'Schmidt', 'Schneider', 'Fischer', 'Weber',
    'Meyer', 'Wagner', 'Becker', 'Schulz', 'Hoffmann',
    'Schäfer', 'Koch', 'Bauer', 'Richter', 'Klein',
    'Wolf', 'Schröder', 'Neumann', 'Schwarz', 'Zimmermann'
  ];
  
  // German first names
  const firstNames = [
    'Alexander', 'Emma', 'Maximilian', 'Sophie', 'Paul',
    'Maria', 'Thomas', 'Anna', 'Michael', 'Laura',
    'Daniel', 'Julia', 'Andreas', 'Sarah', 'Stefan'
  ];
  
  const generateRandomPerson = () => {
    const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
    const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
    const street = streets[Math.floor(Math.random() * streets.length)];
    const houseNumber = Math.floor(Math.random() * 150) + 1;
    const city = cities[Math.floor(Math.random() * cities.length)];
    const postalCode = Math.floor(Math.random() * 89999) + 10000;
    
    return {
      Name: `${firstName} ${lastName}`,
      Straße: `${street} ${houseNumber}`,
      PLZ: postalCode,
      Ort: city
    };
  };
  
  const generateTestData = () => {
    // Generate 100 addresses
    const addresses = Array(100).fill(null).map(generateRandomPerson);
    
    // Generate 30 bank transactions
    const transactions = Array(30).fill(null).map((_, index) => {
      const date = new Date(2024, 0, Math.floor(Math.random() * 31) + 1);
      const amount = (Math.random() * 990 + 10).toFixed(2);
      
      // 80% of names should match address list
      const useExistingName = Math.random() < 0.8;
      const donor = useExistingName 
        ? addresses[Math.floor(Math.random() * addresses.length)].Name
        : `${firstNames[Math.floor(Math.random() * firstNames.length)]} ${lastNames[Math.floor(Math.random() * lastNames.length)]}`;
      
      return {
        Buchungstag: date.toLocaleDateString('de-DE'),
        'Beguenstigter/Zahlungspflichtiger': donor,
        Betrag: amount,
        Verwendungszweck: `Spende ${date.getFullYear()}`
      };
    });
    
    // Convert to CSV format
    const addressesXLSX = 'Name,Straße,PLZ,Ort\n' + 
      addresses.map(addr => 
        `${addr.Name},${addr.Straße},${addr.PLZ},${addr.Ort}`
      ).join('\n');
    
    const bankCSV = 'Buchungstag,Beguenstigter/Zahlungspflichtiger,Betrag,Verwendungszweck\n' +
      transactions.map(trans => 
        `${trans.Buchungstag},"${trans['Beguenstigter/Zahlungspflichtiger']}",${trans.Betrag},"${trans.Verwendungszweck}"`
      ).join('\n');
    
    setOutput(`=== test_addresses.xlsx ===\n${addressesXLSX}\n\n=== bank_statement.csv ===\n${bankCSV}`);
  };

  return (
    <Card className="w-full">
      <CardHeader>
        <CardTitle>Test Data Generator</CardTitle>
      </CardHeader>
      <CardContent>
        <Button 
          onClick={generateTestData}
          className="mb-4"
        >
          Generate Test Data
        </Button>
        
        <pre className="p-4 bg-gray-100 rounded-lg overflow-auto max-h-96 text-sm">
          {output}
        </pre>
      </CardContent>
    </Card>
  );
};

export default TestDataGenerator;