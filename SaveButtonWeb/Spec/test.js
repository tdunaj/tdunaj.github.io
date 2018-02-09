import Jasmine from 'jasmine';

let jasmine = new Jasmine();
jasmine.loadConfigFile('Spec/support/jasmine.json');
jasmine.execute();