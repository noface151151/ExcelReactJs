import React,{Component} from 'react';
import ReactToPrint from "react-to-print";
import Content from './Content';

class TestPrint extends Component{

    render(){

        return(
            <div>
            <ReactToPrint
              trigger={() => <a href="#">Print this out!</a>}
              content={() => this.componentRef}
            />
            <Content header="Report 1" footer="footer 2" sum="30" ref={el => (this.componentRef = el)} />
          </div>
        )
    }
}

export default TestPrint