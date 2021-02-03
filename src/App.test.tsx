import React from 'react';
import { render, screen } from '@testing-library/react';
import { Pagination } from '@uifabric/experiments';
import { initializeIcons } from '@fluentui/react';
initializeIcons();

test('renders learn react link', async () => {
  render(<Pagination
    pageCount={10}
    itemsPerPage={10}
    totalItemCount={268}
    format={'buttons'}
    previousPageAriaLabel={'previous page'}
    nextPageAriaLabel={'next page'}
    firstPageAriaLabel={'first page'}
    lastPageAriaLabel={'last page'}
    pageAriaLabel={'page'}
    selectedAriaLabel={'selected'}
  />);
  let elem = await screen.getByLabelText("first page");
  expect(elem).toBeInTheDocument();
});
