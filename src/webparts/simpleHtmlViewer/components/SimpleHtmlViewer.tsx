import * as React from 'react';
import { ISimpleHtmlViewerProps } from './ISimpleHtmlViewerProps';

const SimpleHtmlViewer: React.FC<ISimpleHtmlViewerProps> = (props) => {
  const { html = `<div style="height:200px;width:100%;background-color:red;font-color:white">prease input html</div>` } = props;
  const htmlProps = React.useMemo(() => ({__html:html}), [html]);
  return (
    <div className="shv" dangerouslySetInnerHTML={htmlProps}>

    </div>
  );
};
export default SimpleHtmlViewer;
