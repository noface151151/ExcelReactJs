import React, { Component } from "react";

class Content extends Component {
  static defaultProps = {
    header: "Report",
    footer: "Test Report",
    sum: 0
  };

  render() {
    return (
      <div>
        <h3>{this.props.header}</h3>
        <h1>{this.props.sum}</h1>
        <table>
          <thead>
            <th>column 1</th>
            <th>column 2</th>
            <th>column 3</th>
          </thead>
          <tbody>
            <tr>
              <td>data 1</td>
              <td>data 2</td>
              <td>data 3</td>
            </tr>
            <tr>
              <td>data 1</td>
              <td>data 2</td>
              <td>data 3</td>
            </tr>
            <tr>
              <td>data 1</td>
              <td>data 2</td>
              <td>data 3</td>
            </tr>
          </tbody>
        </table>
        <h5>{this.props.footer}</h5>
      </div>
    );
  }
}

export default Content;
