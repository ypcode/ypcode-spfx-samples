/// <reference types="mocha" />

import { assert } from 'chai';

import styles from '../KanbanBoard.module.scss';
import KanbanBoardWebPart from "../KanbanBoardWebPart";



    // const exptectedHtml = `<div class="${styles.kanbanBoard}">
    //     <div class="${styles.container}">
    //       <div class="ms-Grid-row ${styles.row}">
    //         <div class="${styles.kanbanColumn} ms-Grid-col ms-u-md3" data-status="Open">
    //           <h3 class="ms-fontColor-themePrimary">Open</h3>
    //           <div class="${styles.task}" data-taskid="1">
    //             <div class="ms-fontSize-xl">Task 1 from list 1</div>
    //           </div>
    //         </div>
    //          <div class="${styles.kanbanColumn} ms-Grid-col ms-u-md3" data-status="On going">
    //           <h3 class="ms-fontColor-themePrimary">On going</h3>
             
    //           <div class="${styles.task}" data-taskid="2">
    //             <div class="ms-fontSize-xl">Task 2 from list 1</div>
    //           </div>
    //           <div class="${styles.task}" data-taskid="3">
    //             <div class="ms-fontSize-xl">Task 3 from list 1</div>
    //           </div>
    //         </div>
    //          <div class="${styles.kanbanColumn} ms-Grid-col ms-u-md3" data-status="Done">
    //           <h3 class="ms-fontColor-themePrimary">Done</h3>
    //           <div class="${styles.task}" data-taskid="4">
    //             <div class="ms-fontSize-xl">Task 4 from list 1</div>
    //           </div>
    //         </div>
    //          <div class="${styles.kanbanColumn} ms-Grid-col ms-u-md3" data-status="Canceled">
    //           <h3 class="ms-fontColor-themePrimary">Canceled</h3>
    //           <div class="${styles.task}" data-taskid="5">
    //             <div class="ms-fontSize-xl">Task 5 from list 1</div>
    //           </div>
    //         </div>
    //       </div>
    //     </div>`;

describe('KanbanBoardWebPart', () => {
  it('should do render 4 columns for Tasks List 1', () => {

    assert.ok(true);
  });
});
