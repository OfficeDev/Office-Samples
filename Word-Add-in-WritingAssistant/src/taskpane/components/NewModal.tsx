import Button from "react-bootstrap/Button";
import Modal from "react-bootstrap/Modal";
import React from "react";
import { acceptAction, insertAnnotations, rejectAction } from "../office-document";

export interface MyModalProps {
  show: boolean;
  handleClose: () => void;
  eventName: string;
  eventMessage: string;
  annotationId: string;
  paraIds: string[];
}

const MyModal: React.FC<MyModalProps> = (props: MyModalProps) => {
  const handleClick = async (func: (...args: any[]) => any, ...args: any[]) => {
    await func(...args);
    props.handleClose();
  };

  const handleGrammerChecking = async (args: string[]) => {
    await insertAnnotations(args);
    props.handleClose();
  };

  return (
    <>
      <Modal
        show={props.show}
        size="lg"
        aria-labelledby="contained-modal-title-vcenter"
        centered={true}
        keyboard={false}
      >
        <Modal.Header>
          <Modal.Title>{props.eventName}</Modal.Title>
        </Modal.Header>
        <Modal.Body>{props.eventMessage}</Modal.Body>
        <Modal.Footer>
          {props.eventName === "AnnotationHovered" ? (
            <>
              <Button variant="primary" onClick={() => handleClick(() => acceptAction(props.annotationId))}>
                Accept
              </Button>
              <Button variant="danger" onClick={() => handleClick(() => rejectAction(props.annotationId))}>
                Reject
              </Button>
            </>
          ) : (
            <></>
          )}
          {props.eventName === "ParagraphAdded" ? (
            <>
              <Button variant="primary" onClick={() => handleGrammerChecking(props.paraIds)}>
                Check Grammar
              </Button>
            </>
          ) : (
            <></>
          )}
          <Button variant="secondary" onClick={props.handleClose}>
            Cancel
          </Button>
        </Modal.Footer>
      </Modal>
    </>
  );
};

export default MyModal;
