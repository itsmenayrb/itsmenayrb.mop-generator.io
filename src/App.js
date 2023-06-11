import React, { useState } from 'react';
import './App.css';
import { saveAs } from 'file-saver';
import { Document, Packer, Paragraph, TextRun } from 'docx';

function App() {

  const [value, setValue] = useState('');

  const onChange = e => setValue(e.target.value);

  const textRun = (text = '', bold = false, italics = false) => (
    new TextRun({
      text,
      bold,
      italics,
      font: 'Calibri',
      size: 24
    })
  )

  const generateBackupFiles = (apis = [], type = 'ubp') => {
    const backupSection = [];
    if (apis.length > 0) {
      apis.forEach(api => {
        if (api.existing) {
          const section = [
            new Paragraph({
              children: [
                textRun(`${type === 'ubp' ? 'ubp' : 'core'}/${api.apiName.toLowerCase().slice(0, -6).replace(/ /g, '-')}_${api.apiName.toLowerCase().slice(-5)}.yaml`, true),
              ],
              bullet: { level: 0 }
            }),
          ];
          backupSection.push(...section);
        }
      })
    }
    return backupSection;
  }

  const generateRevertProcedure = (hasRevert = true, ubpApis = [], coreApis = []) => {
    const revertSection = [];
    if (hasRevert || ubpApis.length > 0 || coreApis.length > 0) {
      const ubpBackup = generateBackupFiles(ubpApis);
      const coreBackup = generateBackupFiles(coreApis, 'core');
      const section = [
        new Paragraph({
          children: [
            textRun('===Rollback/Revert Procedure', true),
          ],
          spacing: { line: 350, before: 20 * 72 * 0.35, after: 20 * 72 * 0.1 }
        }),
        new Paragraph({
          children: [
            textRun('** NOTE: ', true),
            textRun('For existing APIs, current yaml file needs to be backed up first before uploading the updated yaml file.'),
            textRun( '**', true),
          ],
          spacing: { line: 350, after: 20 * 72 * 0.1 }
        }),
        ...ubpBackup,
        ...coreBackup,
        new Paragraph({
          children: [ textRun('1. If the API is newly created APIs, leave it as is. It will not be triggered by the channels and will not affect existing features.') ],
          spacing: { line: 350, before: 20 * 72 * 0.2 }
        }),
        new Paragraph({
          children: [ textRun('2. If the API is existing, restore backup yaml file.') ]
        }),
      ];
      revertSection.push(...section);
    }
    return revertSection;
  }

  const generateSubscriptions = (subscriptions = []) => {
    const subscriptionSection = [];
    let i = 1;
    if (subscriptions.length > 0) {
      subscriptions.forEach(subscription => {
        if (subscription?.isApplicationExisting) {
          const section = [
            new Paragraph({
              children: [
                textRun(`     ${i++}. Subscribe the application `),
                textRun(`${subscription.applicationName || "Application Name Placeholder 1.0.0"}`, true),
                textRun(subscription?.clientId ? `(${subscription?.clientId})` : "", false, true),
                textRun(' to the following products and send client id and secret to '),
                new TextRun({
                  text: `${subscription.email || 'email@unionbankph.com'}`,
                  font: 'Calibri',
                  color: '#0000FF',
                  underline: {},
                }),
              ],
              spacing: { line: 350, before: 20 * 72 * (i%2 !== 0 && i > 1 ? 0 : 0.2), after: 20 * 72 * 0.05 }
            }),
          ];
          subscriptionSection.push(...section);
        } else {
          const section = [
            new Paragraph({
              children: [
                textRun(`     ${i++}. Create new application named `),
                textRun(`${subscription.applicationName || "Application Name Placeholder 1.0.0"}`, true),
                textRun(' and send client id and secret to '),
                new TextRun({
                  text: `${subscription.email || 'email@unionbankph.com'}`,
                  font: 'Calibri',
                  color: '#0000FF',
                  underline: {},
                }),
              ],
              spacing: { line: 350, before: 20 * 72 * (i%2 !== 0 && i > 1 ? 0 : 0.2), after: 20 * 72 * 0.15 }
            }),
            new Paragraph({
              children: [
                textRun(`     ${i++}. Subscribe new application `),
                textRun(`${subscription.applicationName || "Application Name Placeholder 1.0.0"}`, true),
                textRun(' to the following products:')
              ],
              spacing: { line: 350, after: 20 * 72 * 0.05 }
            }),
          ];
          subscriptionSection.push(...section);
        }

        if (subscription?.products.length > 0) {
          subscription?.products.forEach(product => {
            const section = [
              new Paragraph({
                children: [ textRun(product, true) ],
                bullet: { level: 1 }
              }),
            ];
            subscriptionSection.push(...section);
          })
        } else {
          const section = [
            new Paragraph({
              children: [ textRun("Placeholder Product 1.0.0", true) ],
              bullet: { level: 1 }
            }),
          ];
          subscriptionSection.push(...section);
        }
      })
    } else {
      const section = [
        new Paragraph({
          children: [
            textRun(`     ${i++}. Create new application named `),
            textRun("Application Name Placeholder 1.0.0", true),
            textRun(' and send client id and secret to '),
            new TextRun({
              text: 'email@unionbankph.com',
              font: 'Calibri',
              color: '#0000FF',
              underline: {},
            }),
          ],
          spacing: { line: 350, after: 20 * 72 * 0.15 }
        }),
        new Paragraph({
          children: [
            textRun(`     ${i++}. Subscribe new application `),
            textRun("Application Name Placeholder 1.0.0", true),
            textRun(' to the following products:')
          ],
          spacing: { line: 350, after: 20 * 72 * 0.1 }
        }),
        new Paragraph({
          children: [ textRun("Placeholder Product 1.0.0", true) ],
          bullet: { level: 1 }
        }),
      ];
      subscriptionSection.push(...section);
    }
    return subscriptionSection;
  } 

  const generateProperties = (properties = []) => {
    const propertySection = [];
    if (properties.length > 0) {
      properties.forEach(property => {
        const section = [
          new Paragraph({
            children: [
              textRun(`${property.name}=`),
              textRun(`${property.value || ""}`, true)
            ],
            bullet: { level: 0 },
          }),
        ];
        propertySection.push(...section);
      })
    }
    return propertySection;
  }

  const generateSteps = (api) => {
    const stepsSection = [];
    if (api.existing === false) {
      const properties = generateProperties(api.properties || []);
      const section = [
        new Paragraph({
          children: [
            textRun("1. Create a new API named "),
            textRun(`${api.apiName}`, true),
            textRun(" and upload attached file.")
          ],
        }),
        new Paragraph({
          children: [ textRun("2. Add production properties:") ],
        }),
        ...properties,
        new Paragraph({
          children: [
            textRun("apiLoggerUrl="),
            textRun("'http://172.16.17.149:9000/v1/api/logs'", true)
          ],
          bullet: { level: 0 },
        }),
        new Paragraph({
          children: [
            textRun("3. Create a new product named "),
            textRun(`${api.productName || api.apiName}`, true)
          ],
        }),
        new Paragraph({
          children: [ textRun("4. Add the newly created API to the new product.") ],
        }),
        new Paragraph({
          children: [ textRun("5. Double check if production property values are correctly configured.") ],
        }),
        new Paragraph({
          children: [ textRun("6. Stage and publish to production.") ],
        }),
        new Paragraph({
          children: [
            textRun("7. Make sure that product has "),
            textRun("Published ", true),
            textRun("state in the catalog.")
          ],
          spacing: { line: 350, after: 20 * 72 * 0.2 }
        }),
      ];
      stepsSection.push(...section);
    } else {
      const section = [
        new Paragraph({
          children: [ textRun("1. Download existing API for backup") ],
        }),
        new Paragraph({
          children: [ textRun("2. Update API using the file from the api-file-yamls repo.") ],
        }),
        new Paragraph({
          children: [
            textRun("3. "),
            textRun("COMPARE", true),
            textRun(" and "),
            textRun("COPY", true),
            textRun(" catalog "),
            textRun("production", true),
            textRun(" property values from backup file."),
          ]
        }),
        new Paragraph({
          children: [
            textRun("4. Double check values of catalog properties for "),
            textRun("production", true)
          ],
        }),
        new Paragraph({
          children: [
            textRun("5. Stage and publish product to "),
            textRun("production", true)
          ],
        }),
        new Paragraph({
          children: [
            textRun("6. Double check if the API is in published state for "),
            textRun("production", true),
            textRun(" catalogs"),
          ],
          spacing: { line: 350, after: 20 * 72 * 0.2 }
        }),
      ];
      stepsSection.push(...section);
    }
    return stepsSection;
  }

  const generateAPis = (apis = [], type = 'ubp') => {
    const apiSection = [];
    apis.forEach((api, i) => {
      const steps = generateSteps(api);
      const section = [
        new Paragraph({
          children: [ textRun(type === 'ubp' ? `===UBP(${i + 1}/${apis.length})` : `===CORE(${i + 1}/${apis.length})`, true) ],
        }),
        new Paragraph({
          children: [ textRun(`API: ${api.apiName}`) ],
        }),
        new Paragraph({
          children: [ textRun(`Product: ${api.productName || api.apiName}`) ],
        }),
        new Paragraph({
          children: [
            textRun("File: "),
            textRun(`${type === 'ubp' ? 'ubp' : 'core'}/${api.apiName.toLowerCase().slice(0, -6).replace(/ /g, '-')}_${api.apiName.toLowerCase().slice(-5)}.yaml`, true),
          ],
          spacing: { line: 350, after: 20 * 72 * 0.15 }
        }),
        ...steps
      ];
      apiSection.push(...section);
    })
    return apiSection;
  }

  const generateMop = () => {
    try {
      const jsonData = JSON.parse(value);
      const checkoutTag = jsonData?.checkoutTag || "";
      const ubpApis = generateAPis(jsonData?.ubp || []);
      const coreApis = generateAPis(jsonData?.core || [], 'core');
      const subscriptions = generateSubscriptions(jsonData?.subscriptions || []);
      const revertProcedure = generateRevertProcedure(jsonData?.hasRevertProcedure || true, jsonData?.ubp || [], jsonData?.core || []);

      const projectName = jsonData.projectName || 'ProjectName';
      const words = projectName.split(" ");
      const transformedText = words.map(word => word.charAt(0).toUpperCase() + word.slice(1)).join("");
      const current = new Date();
      const dateToday = `${current.getFullYear()}-${current.getMonth()+1}-${current.getDate()}`;
      const filename = `${jsonData.deploymentDate || dateToday}_${transformedText}`

      const doc = new Document({
        sections: [
          {
            properties: {},
            children: [
              new Paragraph({
                children: [ textRun("Hi ITSG-TSS Team,") ],
                spacing: { line: 350, before: 20 * 72 * 0.3, after: 20 * 72 * 0.15 }
              }),
              new Paragraph({
                children: [ textRun("May we request to deploy the API below in production. Thank you!") ],
                spacing: { line: 350, before: 20 * 72 * 0.15, after: 20 * 72 * 0.3 }
              }),
              new Paragraph({
                children: [ textRun("Update Repository:", true) ],
                spacing: { line: 350, before: 20 * 72 * 0.1, after: 20 * 72 * 0.05 }
              }),
              new Paragraph({
                children: [
                  textRun("$ git clone "),
                  new TextRun({
                    text: "https://omega-gitlab.unionbankph.com/api-team/file-dumps/api-file-yamls",
                    font: 'Calibri',
                    color: '#0000FF',
                    underline: {},
                  }),
                  textRun(" (if not yet cloned)", false, true)
                ]
              }),
              new Paragraph({
                children: [ textRun("$ cd api-file-yamls") ]
              }),
              new Paragraph({
                children: [ textRun("$ git fetch") ]
              }),
              new Paragraph({
                children: [
                  textRun("$ git checkout tags/"),
                  textRun(checkoutTag, true)
                ],
                spacing: { line: 350, after: 20 * 72 * 0.2 }
              }),
              ...coreApis,
              ...ubpApis,
              new Paragraph({
                children: [
                  textRun('===New Applications and Subscriptions', true),
                ],
                spacing: { line: 350, before: 20 * 72 * 0.3, after: 20 * 72 * 0.15 }
              }),
              ...subscriptions,
              ...revertProcedure,
            ],
          },
        ],
      });

      Packer.toBlob(doc).then((blob) => {
        saveAs(blob, `${filename}.docx`);
      });
    } catch (e) {
      alert(e);
    }
  }

  return (
    <div className="App">
      <header className="App-header">
        <h2>MOP Generator</h2>
        <p>Paste your JSON below</p>
        <div className="card">
          <textarea
            value={value}
            onChange={onChange}
            rows={20}
            style={{ width: '100%' }}
          />
          <div className="button-container">
            <button className='custom-button' onClick={generateMop}>Generate MOP</button>
          </div>
        </div>
      </header>
    </div>
  );
}

export default App;
